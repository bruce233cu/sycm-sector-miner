# -*- coding: utf-8 -*-
import asyncio
import os
import re
import subprocess
import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from openpyxl import Workbook
from playwright.async_api import async_playwright

APP_NAME = "生意参谋赛道数据采集器"
APP_DIR = Path.home() / ".sycm_sector_miner"
EXPORT_DIR = APP_DIR / "exports"
APP_DIR.mkdir(parents=True, exist_ok=True)
EXPORT_DIR.mkdir(parents=True, exist_ok=True)


def compact(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def popularity_upper(text):
    s = compact(text).replace(",", "")
    m = re.search(r"([0-9.]+)\s*(万|千)?\s*[~～\-—到]\s*([0-9.]+)\s*(万|千)?", s)
    if m:
        def conv(v, u):
            x = float(v)
            if u == "万":
                x *= 10000
            elif u == "千":
                x *= 1000
            return x
        return conv(m.group(3), m.group(4))
    m = re.search(r"([0-9.]+)\s*(万|千)?", s)
    if not m:
        return None
    x = float(m.group(1))
    if m.group(2) == "万":
        x *= 10000
    elif m.group(2) == "千":
        x *= 1000
    return x


class ExcelStore:
    HEADERS = [
        "采集日期", "来源页码", "来源搜索词", "关联搜索词",
        "搜索人气", "搜索人气同比", "点击率", "点击率同比",
        "支付转化率", "支付转化率同比", "支付买家数", "支付买家数同比",
        "需求供给比", "需求供给比同比", "天猫商品点击占比", "天猫商品点击占比同比"
    ]

    def __init__(self, path):
        self.path = path
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = "原始数据"
        self.ws.append(self.HEADERS)
        self.ws.freeze_panes = "A2"
        self.wb.save(self.path)

    def append(self, rows):
        for row in rows:
            self.ws.append([row.get(h, "") for h in self.HEADERS])
        self.wb.save(self.path)


class Collector:
    def __init__(self, url, main_min, related_min, output_dir, log_cb, status_cb):
        self.url = url.strip()
        self.main_min = main_min
        self.related_min = related_min
        self.output_dir = Path(output_dir)
        self.log_cb = log_cb
        self.status_cb = status_cb
        self.stop_flag = False

    def log(self, msg):
        self.log_cb(f"[{time.strftime('%H:%M:%S')}] {msg}")

    async def find_table(self, page, required):
        tables = page.locator("table")
        for i in range(await tables.count()):
            table = tables.nth(i)
            try:
                text = compact(await table.inner_text())
                if all(x in text for x in required):
                    return table
            except Exception:
                pass
        return None

    async def ranking_items(self, page):
        table = await self.find_table(page, ["搜索人气", "搜索分析"])
        if table is None:
            raise RuntimeError("没有识别到搜索排行表格")
        rows = table.locator("tbody tr")
        result = []
        for i in range(await rows.count()):
            cells = [compact(x) for x in await rows.nth(i).locator("td").all_inner_texts()]
            if len(cells) < 3:
                continue
            keyword = ""
            for idx in [1, 0, 2, 3]:
                if idx >= len(cells):
                    continue
                c = cells[idx]
                if c and c not in {"搜索分析", "商机发现", "趋势"} and not re.fullmatch(r"[\d\s.%~～+\-万千]+", c):
                    keyword = c
                    break
            pop = ""
            for c in cells:
                if re.search(r"\d", c) and ("~" in c or "～" in c or "万" in c or "千" in c):
                    pop = c
                    break
            if keyword and pop:
                result.append((keyword, pop))
        return result

    async def click_analysis(self, page, keyword):
        table = await self.find_table(page, ["搜索人气", "搜索分析"])
        rows = table.locator("tbody tr")
        for i in range(await rows.count()):
            row = rows.nth(i)
            if keyword not in compact(await row.inner_text()):
                continue
            link = row.get_by_text("搜索分析", exact=True)
            if not await link.count():
                continue
            old_url = page.url
            await link.first.click(timeout=6000)
            await asyncio.sleep(1.2)
            if page.url != old_url or await page.get_by_text("相关分析", exact=True).count():
                return True
        return False

    async def set_filters(self, page):
        for text in ["年同比", "30天"]:
            try:
                loc = page.get_by_text(text, exact=True)
                if await loc.count():
                    await loc.last.click(timeout=4000)
                    await asyncio.sleep(0.5)
            except Exception:
                self.log(f"未确认筛选项：{text}")
        await asyncio.sleep(1.0)

    def split_metric(self, text):
        raw = str(text or "").strip()
        parts = [compact(x) for x in re.split(r"[\r\n]+", raw) if compact(x)]
        if len(parts) >= 2:
            return parts[0], parts[-1]
        m = re.match(r"(.+?)\s+([+\-]?\d[\d,.]*%)$", compact(raw))
        if m:
            return m.group(1).strip(), m.group(2).strip()
        return compact(raw), ""

    async def analysis_rows(self, page, source_kw, page_no):
        table = await self.find_table(page, ["搜索人气", "支付买家数", "需求供给比"])
        if table is None:
            raise RuntimeError("没有识别到搜索分析表格")
        rows = table.locator("tbody tr")
        result = []
        for i in range(await rows.count()):
            cells = [str(x).strip() for x in await rows.nth(i).locator("td").all_inner_texts()]
            if len(cells) < 7:
                continue
            kw = compact(cells[0]).replace("搜索词", "").strip()
            metrics = [self.split_metric(cells[j]) for j in range(1, min(7, len(cells)))]
            while len(metrics) < 6:
                metrics.append(("", ""))
            upper = popularity_upper(metrics[0][0])
            if upper is not None and upper < self.related_min:
                break
            result.append({
                "采集日期": time.strftime("%Y-%m-%d"),
                "来源页码": page_no,
                "来源搜索词": source_kw,
                "关联搜索词": kw,
                "搜索人气": metrics[0][0], "搜索人气同比": metrics[0][1],
                "点击率": metrics[1][0], "点击率同比": metrics[1][1],
                "支付转化率": metrics[2][0], "支付转化率同比": metrics[2][1],
                "支付买家数": metrics[3][0], "支付买家数同比": metrics[3][1],
                "需求供给比": metrics[4][0], "需求供给比同比": metrics[4][1],
                "天猫商品点击占比": metrics[5][0], "天猫商品点击占比同比": metrics[5][1],
            })
        return result

    async def next_page(self, page):
        for loc in [page.get_by_text("下一页", exact=True), page.locator('[title="下一页"]'), page.locator('[aria-label="下一页"]')]:
            try:
                if not await loc.count():
                    continue
                el = loc.last
                cls = (await el.get_attribute("class") or "").lower()
                aria = (await el.get_attribute("aria-disabled") or "").lower()
                if "disabled" in cls or aria == "true":
                    return False
                await el.click(timeout=5000)
                await asyncio.sleep(1.0)
                return True
            except Exception:
                pass
        return False

    async def run(self):
        self.output_dir.mkdir(parents=True, exist_ok=True)
        out = self.output_dir / f"生意参谋_搜索分析_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        store = ExcelStore(out)
        async with async_playwright() as p:
            browser = await p.chromium.connect_over_cdp("http://127.0.0.1:9222")
            if not browser.contexts:
                raise RuntimeError("Chrome 已连接，但没有可用页面")
            context = browser.contexts[0]
            pages = context.pages
            page = next((x for x in pages if "sycm.taobao.com" in x.url), pages[-1] if pages else await context.new_page())
            if self.url:
                await page.goto(self.url, wait_until="domcontentloaded", timeout=30000)
                await asyncio.sleep(1)
            page_no = 1
            processed = 0
            while not self.stop_flag:
                items = await self.ranking_items(page)
                if not items:
                    raise RuntimeError("当前页没有识别到关键词")
                for keyword, pop in items:
                    if self.stop_flag:
                        break
                    upper = popularity_upper(pop)
                    if upper is not None and upper < self.main_min:
                        self.log(f"图1搜索人气低于 {self.main_min}，停止后续主词")
                        return str(out)
                    processed += 1
                    self.status_cb(f"处理中：{keyword}｜已处理 {processed} 个主词")
                    self.log(f"主词：{keyword}（{pop}）")
                    source_url = page.url
                    if not await self.click_analysis(page, keyword):
                        self.log("未进入搜索分析，跳过")
                        continue
                    await self.set_filters(page)
                    rows = await self.analysis_rows(page, keyword, page_no)
                    store.append(rows)
                    self.log(f"保存 {len(rows)} 条关联词")
                    try:
                        await page.go_back(wait_until="domcontentloaded", timeout=12000)
                        await asyncio.sleep(1)
                    except Exception:
                        await page.goto(source_url, wait_until="domcontentloaded", timeout=20000)
                        await asyncio.sleep(1)
                if self.stop_flag:
                    break
                if not await self.next_page(page):
                    break
                page_no += 1
        return str(out)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("980x700")
        self.worker = None
        self.collector = None

        frame = ttk.Frame(root, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=APP_NAME, font=("Microsoft YaHei UI", 18, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 14))

        ttk.Label(frame, text="搜索排行 URL（Chrome 已停在图1时可留空）").grid(row=1, column=0, sticky="w")
        self.url = tk.StringVar()
        ttk.Entry(frame, textvariable=self.url, width=90).grid(row=1, column=1, columnspan=3, sticky="ew", pady=5)

        ttk.Label(frame, text="图1主词最低搜索人气").grid(row=2, column=0, sticky="w")
        self.main_min = tk.IntVar(value=300)
        ttk.Entry(frame, textvariable=self.main_min, width=12).grid(row=2, column=1, sticky="w")

        ttk.Label(frame, text="图2关联词最低搜索人气").grid(row=2, column=2, sticky="w")
        self.related_min = tk.IntVar(value=100)
        ttk.Entry(frame, textvariable=self.related_min, width=12).grid(row=2, column=3, sticky="w")

        ttk.Label(frame, text="导出目录").grid(row=3, column=0, sticky="w")
        self.output_dir = tk.StringVar(value=str(EXPORT_DIR))
        ttk.Entry(frame, textvariable=self.output_dir).grid(row=3, column=1, columnspan=2, sticky="ew", pady=5)
        ttk.Button(frame, text="选择目录", command=self.choose_dir).grid(row=3, column=3, sticky="w")

        buttons = ttk.Frame(frame)
        buttons.grid(row=4, column=0, columnspan=4, sticky="w", pady=(12, 8))
        ttk.Button(buttons, text="启动采集专用 Chrome", command=self.launch_chrome).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="测试 Chrome 连接", command=self.test_connection).pack(side="left", padx=4)
        ttk.Button(buttons, text="开始采集", command=self.start).pack(side="left", padx=4)
        ttk.Button(buttons, text="停止", command=self.stop).pack(side="left", padx=4)
        ttk.Button(buttons, text="打开导出目录", command=self.open_output).pack(side="left", padx=10)

        self.status = tk.StringVar(value="首次使用：启动采集专用 Chrome，并在其中手动登录一次生意参谋")
        ttk.Label(frame, textvariable=self.status).grid(row=5, column=0, columnspan=4, sticky="w")

        ttk.Label(frame, text="运行日志").grid(row=6, column=0, columnspan=4, sticky="w", pady=(10, 2))
        self.log_box = tk.Text(frame, height=26, wrap="word")
        self.log_box.grid(row=7, column=0, columnspan=4, sticky="nsew")

        frame.rowconfigure(7, weight=1)
        for i in range(4):
            frame.columnconfigure(i, weight=1)

    def log(self, msg):
        self.root.after(0, lambda: (self.log_box.insert("end", msg + "\n"), self.log_box.see("end")))

    def set_status(self, msg):
        self.root.after(0, lambda: self.status.set(msg))

    def choose_dir(self):
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
        chrome = next((p for p in candidates if p.exists()), None)
        if chrome is None:
            messagebox.showerror("未找到 Chrome", "请先安装 Google Chrome")
            return
        profile = Path(os.environ.get("LocalAppData", str(Path.home()))) / "SYCMCollectorChrome"
        profile.mkdir(parents=True, exist_ok=True)
        subprocess.Popen([
            str(chrome), "--remote-debugging-port=9222", "--remote-allow-origins=*",
            f"--user-data-dir={profile}", "https://sycm.taobao.com/"
        ])
        self.set_status("采集专用 Chrome 已启动；首次使用请手动登录生意参谋")

    def test_connection(self):
        def worker():
            async def test():
                async with async_playwright() as p:
                    browser = await p.chromium.connect_over_cdp("http://127.0.0.1:9222")
                    return [pg.url for c in browser.contexts for pg in c.pages]
            try:
                urls = asyncio.run(test())
                self.log("Chrome 连接成功：" + " | ".join(urls[:5]))
                self.set_status("Chrome 连接成功")
            except Exception as e:
                self.log("Chrome 连接失败：" + str(e))
                self.set_status("Chrome 连接失败")
        threading.Thread(target=worker, daemon=True).start()

    def start(self):
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("提示", "任务正在运行")
            return
        self.collector = Collector(
            self.url.get(), int(self.main_min.get()), int(self.related_min.get()),
            self.output_dir.get(), self.log, self.set_status
        )

        def worker():
            try:
                result = asyncio.run(self.collector.run())
                self.set_status("完成：" + result)
                self.root.after(0, lambda: messagebox.showinfo("完成", "采集完成：\n" + result))
            except Exception as e:
                err = str(e)
                self.log("运行失败：" + err)
                self.set_status("运行失败")
                self.root.after(0, lambda msg=err: messagebox.showerror("运行失败", msg))

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    def stop(self):
        if self.collector:
            self.collector.stop_flag = True
            self.set_status("正在停止")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
