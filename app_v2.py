# -*- coding: utf-8 -*-
import asyncio, json, os, random, re, subprocess, sys, threading, time
from datetime import date, timedelta
from pathlib import Path
from urllib.parse import urlparse, parse_qs
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from openpyxl import Workbook
from playwright.async_api import async_playwright

APP = "生意参谋赛道数据采集器 API V2"
ROOT = Path.home() / ".sycm_sector_miner_v2"
OUT = ROOT / "exports"
DBG = ROOT / "debug"
for p in (ROOT, OUT, DBG):
    p.mkdir(parents=True, exist_ok=True)

RANK_ENDPOINT = "/mc/mq/mkt/keyword/rank/pro.json"
RELATED_ENDPOINT = "/mc/mq/mkt/keyword/relate/analysis.json"
RANK_REFERER = "/mc/free/search_rank"
RELATED_REFERER = "/mc/free/search_analysis"


def compact(v):
    return re.sub(r"\s+", " ", str("" if v is None else v)).strip()


def metric(v):
    return v.get("value", "") if isinstance(v, dict) else ("" if v is None else v)


def comparison(v):
    return v.get("cycleCrc", "") if isinstance(v, dict) else ""


def popularity_upper(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = compact(v).replace(",", "")

    def cv(x, unit):
        n = float(x)
        if unit == "万":
            return n * 10000
        if unit == "千":
            return n * 1000
        return n

    m = re.search(r"([0-9.]+)\s*(万|千)?\s*[~～\-—到]\s*([0-9.]+)\s*(万|千)?", s)
    if m:
        return cv(m.group(3), m.group(4))
    m = re.search(r"([0-9.]+)\s*(万|千)?", s)
    return cv(m.group(1), m.group(2)) if m else None


def window_30_days(today=None):
    today = today or date.today()
    end = today - timedelta(days=1)
    start = end - timedelta(days=29)
    return f"{start.isoformat()}|{end.isoformat()}"


def normalize_url(value):
    s = compact(value)
    if s.startswith("http://") or s.startswith("https://"):
        return s
    if s.startswith("/mc/free/search_rank"):
        return "https://sycm.taobao.com" + s
    if s.startswith("/search_rank"):
        return "https://sycm.taobao.com/mc/free" + s
    if s.startswith("search_rank"):
        return "https://sycm.taobao.com/mc/free/" + s
    return s


def parse_context(url):
    q = parse_qs(urlparse(url).query)
    get = lambda key: (q.get(key) or [""])[0]
    return {
        "parentCateId": get("parentCateId"),
        "cateId": get("cateId"),
        "cateFlag": get("cateFlag"),
    }


def build_rank_params(date_range, context, page=1):
    return {
        "dateRange": date_range,
        "dateType": "day",
        "page": page,
        "pageSize": 50,
        "order": "desc",
        "marketVersion": "free",
        "parentCateId": context.get("parentCateId", ""),
        "cateId": context.get("cateId", ""),
        "cateFlag": context.get("cateFlag", ""),
        "kwType": "search",
        "rankType": "hot",
        "orderBy": "seIpvUvHits",
        "device": 0,
    }


def build_related_params(date_range, keyword, page=1):
    return {
        "dateRange": date_range,
        "dateType": "day",
        "page": page,
        "pageSize": 100,
        "order": "desc",
        "marketVersion": "free",
        "keyWord": keyword,
        "rankType": "related",
        "cycleFlag": "yearSync",
        "orderBy": "seIpvUvHits",
    }


class ExcelStore:
    HEADERS = [
        "采集日期", "数据周期", "对比方式", "主词页码", "主词排名",
        "来源搜索词", "主词搜索人气", "关联搜索词",
        "搜索人气", "搜索人气同比", "点击率", "点击率同比",
        "支付转化率", "支付转化率同比", "支付买家数", "支付买家数同比",
        "需求供给比", "需求供给比同比", "天猫商品点击占比", "天猫商品点击占比同比",
    ]

    def __init__(self, path):
        self.path = Path(path)
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = "原始数据"
        self.ws.append(self.HEADERS)
        self.ws.freeze_panes = "A2"
        self.wb.save(self.path)

    def add(self, rows):
        for row in rows:
            self.ws.append([row.get(h, "") for h in self.HEADERS])
        self.wb.save(self.path)


class Collector:
    def __init__(self, url, main_min, related_min, output_dir, log_cb, status_cb):
        self.url = normalize_url(url)
        self.main_min = int(main_min)
        self.related_min = int(related_min)
        self.output_dir = Path(output_dir)
        self.log_cb = log_cb
        self.status_cb = status_cb
        self.stop_flag = False
        self.date_range = window_30_days()

    def log(self, msg):
        self.log_cb(f"[{time.strftime('%H:%M:%S')}] {msg}")

    async def api(self, page, endpoint, params, referer):
        js = """
        async a => {
          const url = new URL(a.endpoint, location.origin);
          Object.entries(a.params).forEach(([k,v]) => url.searchParams.set(k, String(v)));
          try {
            const response = await fetch(url, {
              credentials: 'include',
              headers: {
                accept: '*/*',
                'sycm-referer': a.referer,
                'x-sycm-collector': '1'
              }
            });
            const text = await response.text();
            let payload = null;
            try { payload = JSON.parse(text); } catch (_) {}
            return {
              status: response.status,
              payload,
              text: payload ? null : text.slice(0, 500),
              url: url.toString()
            };
          } catch (e) {
            return {status: 0, payload: null, text: String(e), url: url.toString()};
          }
        }
        """
        last = None
        for attempt in range(3):
            if self.stop_flag:
                raise RuntimeError("已停止")
            last = await page.evaluate(js, {
                "endpoint": endpoint,
                "params": params,
                "referer": referer,
            })
            payload = last.get("payload") if isinstance(last, dict) else None
            if isinstance(payload, dict) and int(payload.get("code", -1)) == 0:
                data = payload.get("data") or {}
                rows = list(data.get("data") or [])
                total = int(data.get("recordCount") or 0)
                return rows, total
            if attempt < 2:
                self.log(f"接口失败，第 {attempt + 1} 次重试")
                await asyncio.sleep(1.5)

        debug_file = DBG / f"api_error_{int(time.time())}.json"
        debug_file.write_text(json.dumps(last, ensure_ascii=False, indent=2), encoding="utf-8")
        status = last.get("status") if isinstance(last, dict) else None
        if status in (401, 403):
            raise RuntimeError("登录态已失效，请重新登录生意参谋")
        msg = ""
        if isinstance(last, dict) and isinstance(last.get("payload"), dict):
            msg = compact(last["payload"].get("message"))
        suffix = f"：{msg}" if msg else ""
        raise RuntimeError(f"生意参谋接口返回失败{suffix}。调试文件：{debug_file}")

    async def preflight(self, page, context):
        self.status_cb("正在做接口自检…")
        self.log("开始接口自检：图1搜索排行")
        rank_rows, _ = await self.api(
            page,
            RANK_ENDPOINT,
            build_rank_params(self.date_range, context, 1),
            RANK_REFERER,
        )
        if not rank_rows:
            raise RuntimeError("接口自检失败：图1搜索排行返回为空")
        first = next((x for x in rank_rows if isinstance(x, dict) and compact(metric(x.get("searchWord")))), None)
        if first is None:
            raise RuntimeError("接口自检失败：图1缺少 searchWord 字段")
        keyword = compact(metric(first.get("searchWord")))
        if "seIpvUvHits" not in first:
            raise RuntimeError("接口自检失败：图1缺少 seIpvUvHits 字段")

        self.log(f"图1自检通过，测试主词：{keyword}")
        self.log("开始接口自检：图2相关搜索词")
        related_rows, _ = await self.api(
            page,
            RELATED_ENDPOINT,
            build_related_params(self.date_range, keyword, 1),
            RELATED_REFERER,
        )
        if related_rows:
            sample = next((x for x in related_rows if isinstance(x, dict)), None)
            if sample is None or "relatedSekeyword" not in sample:
                raise RuntimeError("接口自检失败：图2缺少 relatedSekeyword 字段")
        self.log(f"图2自检通过，返回 {len(related_rows)} 条")
        self.status_cb("接口自检通过，开始正式采集")

    async def collect_related(self, page, keyword, main_page, main_rank, main_popularity):
        result = []
        page_no = 1
        while not self.stop_flag and page_no <= 100:
            params = build_related_params(self.date_range, keyword, page_no)
            rows, total = await self.api(page, RELATED_ENDPOINT, params, RELATED_REFERER)
            if not rows:
                break
            threshold_hit = False
            for item in rows:
                if not isinstance(item, dict):
                    continue
                pop = metric(item.get("seIpvUvHits"))
                up = popularity_upper(pop)
                if up is not None and up < self.related_min:
                    self.log(f"图2达到阈值：{keyword} → {pop} < {self.related_min}")
                    threshold_hit = True
                    break
                result.append({
                    "采集日期": time.strftime("%Y-%m-%d"),
                    "数据周期": self.date_range,
                    "对比方式": "年同比",
                    "主词页码": main_page,
                    "主词排名": main_rank,
                    "来源搜索词": keyword,
                    "主词搜索人气": main_popularity,
                    "关联搜索词": metric(item.get("relatedSekeyword")),
                    "搜索人气": pop,
                    "搜索人气同比": comparison(item.get("seIpvUvHits")),
                    "点击率": metric(item.get("freeClkRate")),
                    "点击率同比": comparison(item.get("freeClkRate")),
                    "支付转化率": metric(item.get("payConvRate")),
                    "支付转化率同比": comparison(item.get("payConvRate")),
                    "支付买家数": metric(item.get("payByrCnt")),
                    "支付买家数同比": comparison(item.get("payByrCnt")),
                    "需求供给比": metric(item.get("simWeight")),
                    "需求供给比同比": comparison(item.get("simWeight")),
                    "天猫商品点击占比": metric(item.get("tmaoClickRatio")),
                    "天猫商品点击占比同比": comparison(item.get("tmaoClickRatio")),
                })
            if threshold_hit or len(rows) < 100 or (total and page_no * 100 >= total):
                break
            page_no += 1
            await asyncio.sleep(random.uniform(0.5, 0.9))
        return result

    async def run(self):
        self.output_dir.mkdir(parents=True, exist_ok=True)
        output_path = self.output_dir / f"生意参谋_API_V2_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        store = ExcelStore(output_path)

        async with async_playwright() as p:
            browser = await p.chromium.connect_over_cdp("http://127.0.0.1:9222")
            if not browser.contexts:
                raise RuntimeError("没有连接到采集专用 Chrome")
            context_obj = browser.contexts[0]
            pages = context_obj.pages
            page = next((x for x in pages if "sycm.taobao.com" in x.url), pages[-1] if pages else await context_obj.new_page())

            if self.url:
                await page.goto(self.url, wait_until="domcontentloaded", timeout=30000)
                await asyncio.sleep(1)
            if "sycm.taobao.com" not in page.url:
                raise RuntimeError("请先在采集专用 Chrome 中登录生意参谋")

            context = parse_context(self.url or page.url)
            if not context["cateId"]:
                context = parse_context(page.url)
            if not context["cateId"]:
                raise RuntimeError("没有识别到 cateId，请粘贴图1完整网址")

            self.log(f"类目 cateId={context['cateId']}｜过去30天={self.date_range}｜年同比")
            await self.preflight(page, context)

            main_page = 1
            rank = 0
            processed = 0
            while not self.stop_flag and main_page <= 100:
                rank_rows, total = await self.api(
                    page,
                    RANK_ENDPOINT,
                    build_rank_params(self.date_range, context, main_page),
                    RANK_REFERER,
                )
                if not rank_rows:
                    break
                global_stop = False
                for item in rank_rows:
                    if self.stop_flag:
                        break
                    if not isinstance(item, dict):
                        continue
                    rank += 1
                    keyword = compact(metric(item.get("searchWord")))
                    pop = metric(item.get("seIpvUvHits"))
                    if not keyword:
                        continue
                    up = popularity_upper(pop)
                    if up is not None and up < self.main_min:
                        self.log(f"图1达到阈值：{keyword} {pop} < {self.main_min}，停止后续主词")
                        global_stop = True
                        break
                    processed += 1
                    self.status_cb(f"处理中：{keyword}｜主词 {processed}｜排名 {rank}")
                    self.log(f"主词 #{rank}：{keyword}（{pop}）")
                    rows = await self.collect_related(page, keyword, main_page, rank, pop)
                    store.add(rows)
                    self.log(f"已保存 {len(rows)} 条关联词")
                    await asyncio.sleep(random.uniform(0.7, 1.2))

                if global_stop or self.stop_flag or len(rank_rows) < 50 or (total and main_page * 50 >= total):
                    break
                main_page += 1
                await asyncio.sleep(random.uniform(0.5, 0.9))

        self.log(f"完成：{output_path}")
        return str(output_path)


class App:
    def __init__(self, root):
        self.root = root
        root.title(APP)
        root.geometry("1020x740")
        self.worker = None
        self.collector = None

        f = ttk.Frame(root, padding=16)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text=APP, font=("Microsoft YaHei UI", 18, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 14))

        ttk.Label(f, text="搜索排行 URL").grid(row=1, column=0, sticky="w")
        self.url = tk.StringVar()
        ttk.Entry(f, textvariable=self.url, width=92).grid(row=1, column=1, columnspan=3, sticky="ew", pady=5)

        ttk.Label(f, text="图1主词最低搜索人气").grid(row=2, column=0, sticky="w")
        self.main_min = tk.IntVar(value=300)
        ttk.Entry(f, textvariable=self.main_min, width=12).grid(row=2, column=1, sticky="w")
        ttk.Label(f, text="图2关联词最低搜索人气").grid(row=2, column=2, sticky="w")
        self.related_min = tk.IntVar(value=100)
        ttk.Entry(f, textvariable=self.related_min, width=12).grid(row=2, column=3, sticky="w")

        ttk.Label(f, text="数据周期").grid(row=3, column=0, sticky="w")
        ttk.Label(f, text="过去30个完整自然日").grid(row=3, column=1, sticky="w")
        ttk.Label(f, text="对比方式").grid(row=3, column=2, sticky="w")
        ttk.Label(f, text="年同比").grid(row=3, column=3, sticky="w")

        ttk.Label(f, text="导出目录").grid(row=4, column=0, sticky="w")
        self.output_dir = tk.StringVar(value=str(OUT))
        ttk.Entry(f, textvariable=self.output_dir).grid(row=4, column=1, columnspan=2, sticky="ew", pady=5)
        ttk.Button(f, text="选择目录", command=self.pick_dir).grid(row=4, column=3, sticky="w")

        buttons = ttk.Frame(f)
        buttons.grid(row=5, column=0, columnspan=4, sticky="w", pady=(12, 8))
        for title, command in [
            ("启动采集专用 Chrome", self.launch_chrome),
            ("测试 Chrome", self.test_chrome),
            ("开始采集", self.start),
            ("停止", self.stop),
            ("打开导出目录", self.open_output),
        ]:
            ttk.Button(buttons, text=title, command=command).pack(side="left", padx=4)

        self.status = tk.StringVar(value="V2：开始前自动验证图1和图2接口")
        ttk.Label(f, textvariable=self.status).grid(row=6, column=0, columnspan=4, sticky="w")
        ttk.Label(f, text="运行日志").grid(row=7, column=0, columnspan=4, sticky="w", pady=(10, 2))
        self.log_box = tk.Text(f, height=26, wrap="word")
        self.log_box.grid(row=8, column=0, columnspan=4, sticky="nsew")
        ttk.Label(f, text="首次：启动专用 Chrome → 登录生意参谋 → 粘贴图1完整网址 → 开始采集。", foreground="#555").grid(row=9, column=0, columnspan=4, sticky="w", pady=(10, 0))

        f.rowconfigure(8, weight=1)
        for i in range(4):
            f.columnconfigure(i, weight=1)

    def log(self, msg):
        self.root.after(0, lambda x=msg: (self.log_box.insert("end", x + "\n"), self.log_box.see("end")))

    def set_status(self, msg):
        self.root.after(0, lambda x=msg: self.status.set(x))

    def pick_dir(self):
        path = filedialog.askdirectory(initialdir=self.output_dir.get())
        if path:
            self.output_dir.set(path)

    def open_output(self):
        Path(self.output_dir.get()).mkdir(parents=True, exist_ok=True)
        os.startfile(self.output_dir.get())

    def launch_chrome(self):
        candidates = [
            Path(os.environ.get("ProgramFiles", "")) / "Google/Chrome/Application/chrome.exe",
            Path(os.environ.get("ProgramFiles(x86)", "")) / "Google/Chrome/Application/chrome.exe",
            Path(os.environ.get("LocalAppData", "")) / "Google/Chrome/Application/chrome.exe",
        ]
        chrome = next((x for x in candidates if x.exists()), None)
        if not chrome:
            messagebox.showerror("未找到 Chrome", "请先安装 Google Chrome")
            return
        profile = ROOT / "chrome-profile"
        profile.mkdir(parents=True, exist_ok=True)
        subprocess.Popen([
            str(chrome),
            "--remote-debugging-port=9222",
            "--remote-allow-origins=*",
            f"--user-data-dir={profile}",
            "https://sycm.taobao.com/",
        ])
        self.status.set("已启动专用 Chrome；首次请手动登录")

    def test_chrome(self):
        def worker():
            async def run_test():
                async with async_playwright() as p:
                    b = await p.chromium.connect_over_cdp("http://127.0.0.1:9222")
                    return [pg.url for c in b.contexts for pg in c.pages]
            try:
                urls = asyncio.run(run_test())
                self.log("Chrome 连接成功：" + " | ".join(urls[:4]))
                self.set_status("Chrome 连接成功")
                self.root.after(0, lambda: messagebox.showinfo("成功", "Chrome 已连接"))
            except Exception as e:
                err = str(e)
                self.log("Chrome 连接失败：" + err)
                self.root.after(0, lambda x=err: messagebox.showerror("失败", "请先启动采集专用 Chrome。\n\n" + x))
        threading.Thread(target=worker, daemon=True).start()

    def start(self):
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("提示", "任务正在运行")
            return
        if not compact(self.url.get()):
            messagebox.showwarning("缺少 URL", "请粘贴图1完整网址")
            return
        self.collector = Collector(
            self.url.get(),
            self.main_min.get(),
            self.related_min.get(),
            self.output_dir.get(),
            self.log,
            self.set_status,
        )
        self.worker = threading.Thread(target=self._run, daemon=True)
        self.worker.start()

    def _run(self):
        try:
            path = asyncio.run(self.collector.run())
            self.set_status("完成：" + path)
            self.root.after(0, lambda x=path: messagebox.showinfo("完成", "文件已保存：\n" + x))
        except Exception as e:
            err = str(e)
            self.log("运行失败：" + err)
            self.set_status("运行失败")
            self.root.after(0, lambda x=err: messagebox.showerror("运行失败", x))

    def stop(self):
        if self.collector:
            self.collector.stop_flag = True
            self.set_status("正在停止…")


def selftest():
    assert popularity_upper("300 ~ 600") == 600
    assert popularity_upper("1万 ~ 2万") == 20000
    assert popularity_upper(123) == 123
    assert normalize_url("/search_rank?a=1") == "https://sycm.taobao.com/mc/free/search_rank?a=1"
    assert parse_context("https://sycm.taobao.com/mc/free/search_rank?cateId=2")["cateId"] == "2"
    assert window_30_days(date(2026, 9, 15)) == "2026-08-16|2026-09-14"
    rp = build_rank_params("2026-08-16|2026-09-14", {"parentCateId":"1","cateId":"2","cateFlag":"3"}, 1)
    assert rp["pageSize"] == 50 and rp["rankType"] == "hot" and rp["orderBy"] == "seIpvUvHits"
    ap = build_related_params("2026-08-16|2026-09-14", "加热发帽", 1)
    assert ap["pageSize"] == 100 and ap["cycleFlag"] == "yearSync" and ap["rankType"] == "related"
    assert RANK_REFERER == "/mc/free/search_rank"
    assert RELATED_REFERER == "/mc/free/search_analysis"
    print("SELF_TEST_V2_OK")


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        selftest()
    else:
        root = tk.Tk()
        App(root)
        root.mainloop()
