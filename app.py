# -*- coding: utf-8 -*-
import asyncio, json, os, random, re, subprocess, sys, threading, time
from datetime import date,timedelta
from pathlib import Path
from urllib.parse import urlparse,parse_qs
import tkinter as tk
from tkinter import ttk,filedialog,messagebox
from openpyxl import Workbook
from playwright.async_api import async_playwright

APP="生意参谋赛道数据采集器 API版"
ROOT=Path.home()/".sycm_sector_miner"; OUT=ROOT/"exports"; DBG=ROOT/"debug"
for p in (ROOT,OUT,DBG): p.mkdir(parents=True,exist_ok=True)
RANK="/mc/mq/mkt/keyword/rank/pro.json"; RELATED="/mc/mq/mkt/keyword/relate/analysis.json"

def compact(v): return re.sub(r"\s+"," ",str("" if v is None else v)).strip()
def metric(v): return v.get("value","") if isinstance(v,dict) else ("" if v is None else v)
def comp(v): return v.get("cycleCrc","") if isinstance(v,dict) else ""
def upper(v):
    if isinstance(v,(int,float)): return float(v)
    s=compact(v).replace(",","")
    def cv(x,u):
        n=float(x)
        return n*10000 if u=="万" else n*1000 if u=="千" else n
    m=re.search(r"([0-9.]+)\s*(万|千)?\s*[~～\-—到]\s*([0-9.]+)\s*(万|千)?",s)
    if m:return cv(m.group(3),m.group(4))
    m=re.search(r"([0-9.]+)\s*(万|千)?",s)
    return cv(m.group(1),m.group(2)) if m else None

def window30():
    e=date.today()-timedelta(days=1); s=e-timedelta(days=29)
    return f"{s.isoformat()}|{e.isoformat()}"

def norm_url(s):
    s=compact(s)
    if s.startswith("http"): return s
    if s.startswith("/mc/free/search_rank"): return "https://sycm.taobao.com"+s
    if s.startswith("/search_rank"): return "https://sycm.taobao.com/mc/free"+s
    if s.startswith("search_rank"): return "https://sycm.taobao.com/mc/free/"+s
    return s

def ctx(url):
    q=parse_qs(urlparse(url).query)
    g=lambda k:(q.get(k)or[""])[0]
    return {"parentCateId":g("parentCateId"),"cateId":g("cateId"),"cateFlag":g("cateFlag")}

class Store:
    H=["采集日期","数据周期","对比方式","主词页码","主词排名","来源搜索词","主词搜索人气","关联搜索词","搜索人气","搜索人气同比","点击率","点击率同比","支付转化率","支付转化率同比","支付买家数","支付买家数同比","需求供给比","需求供给比同比","天猫商品点击占比","天猫商品点击占比同比"]
    def __init__(self,p):
        self.p=Path(p); self.wb=Workbook(); self.ws=self.wb.active; self.ws.title="原始数据"; self.ws.append(self.H); self.ws.freeze_panes="A2"; self.wb.save(self.p)
    def add(self,rows):
        for r in rows:self.ws.append([r.get(h,"") for h in self.H])
        self.wb.save(self.p)

class Collector:
    def __init__(self,url,mn,rn,out,log,status):
        self.url=norm_url(url); self.mn=int(mn); self.rn=int(rn); self.out=Path(out); self.log_cb=log; self.status=status; self.stop=False; self.dr=window30()
    def log(self,s): self.log_cb(f"[{time.strftime('%H:%M:%S')}] {s}")
    async def api(self,page,endpoint,params):
        js="""async a=>{const u=new URL(a.e,location.origin);Object.entries(a.p).forEach(([k,v])=>u.searchParams.set(k,String(v)));try{const r=await fetch(u,{credentials:'include',headers:{accept:'*/*','sycm-referer':location.pathname,'x-sycm-collector':'1'}});const t=await r.text();let x=null;try{x=JSON.parse(t)}catch(e){};return{status:r.status,json:x,text:x?null:t.slice(0,400),url:u.toString()}}catch(e){return{status:0,json:null,text:String(e)}}}"""
        last=None
        for i in range(3):
            if self.stop: raise RuntimeError("已停止")
            last=await page.evaluate(js,{"e":endpoint,"p":params}); x=last.get("json") if isinstance(last,dict) else None
            if isinstance(x,dict) and int(x.get("code",-1))==0:
                d=x.get("data") or {}; return list(d.get("data") or []),int(d.get("recordCount") or 0)
            if i<2: self.log(f"接口失败，第{i+1}次重试"); await asyncio.sleep(2)
        f=DBG/f"api_error_{int(time.time())}.json"; f.write_text(json.dumps(last,ensure_ascii=False,indent=2),encoding="utf-8")
        if isinstance(last,dict) and last.get("status") in (401,403): raise RuntimeError("登录态已失效，请重新登录生意参谋")
        msg=(last.get("json")or{}).get("message","") if isinstance(last,dict) and isinstance(last.get("json"),dict) else ""
        raise RuntimeError(f"生意参谋接口返回失败{('：'+msg) if msg else ''}。调试文件：{f}")
    async def related(self,page,kw,mp,mr,mpop):
        out=[]; pn=1
        while not self.stop and pn<=100:
            p={"dateRange":self.dr,"dateType":"day","page":pn,"pageSize":100,"order":"desc","marketVersion":"free","keyWord":kw,"rankType":"related","cycleFlag":"yearSync","orderBy":"seIpvUvHits"}
            rows,total=await self.api(page,RELATED,p)
            if not rows: break
            hit=False
            for it in rows:
                if not isinstance(it,dict): continue
                pop=metric(it.get("seIpvUvHits")); u=upper(pop)
                if u is not None and u<self.rn: self.log(f"图2达到阈值：{kw} → {pop} < {self.rn}"); hit=True; break
                out.append({"采集日期":time.strftime("%Y-%m-%d"),"数据周期":self.dr,"对比方式":"年同比","主词页码":mp,"主词排名":mr,"来源搜索词":kw,"主词搜索人气":mpop,"关联搜索词":metric(it.get("relatedSekeyword")),"搜索人气":pop,"搜索人气同比":comp(it.get("seIpvUvHits")),"点击率":metric(it.get("freeClkRate")),"点击率同比":comp(it.get("freeClkRate")),"支付转化率":metric(it.get("payConvRate")),"支付转化率同比":comp(it.get("payConvRate")),"支付买家数":metric(it.get("payByrCnt")),"支付买家数同比":comp(it.get("payByrCnt")),"需求供给比":metric(it.get("simWeight")),"需求供给比同比":comp(it.get("simWeight")),"天猫商品点击占比":metric(it.get("tmaoClickRatio")),"天猫商品点击占比同比":comp(it.get("tmaoClickRatio"))})
            if hit or len(rows)<100 or (total and pn*100>=total): break
            pn+=1; await asyncio.sleep(random.uniform(.5,.9))
        return out
    async def run(self):
        self.out.mkdir(parents=True,exist_ok=True); path=self.out/f"生意参谋_API采集_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"; store=Store(path)
        async with async_playwright() as p:
            b=await p.chromium.connect_over_cdp("http://127.0.0.1:9222")
            if not b.contexts: raise RuntimeError("没有连接到采集专用 Chrome")
            c=b.contexts[0]; pages=c.pages; page=next((x for x in pages if "sycm.taobao.com" in x.url),pages[-1] if pages else await c.new_page())
            if self.url: await page.goto(self.url,wait_until="domcontentloaded",timeout=30000); await asyncio.sleep(1)
            if "sycm.taobao.com" not in page.url: raise RuntimeError("请先在采集专用 Chrome 中登录生意参谋")
            cc=ctx(self.url or page.url)
            if not cc["cateId"]: cc=ctx(page.url)
            if not cc["cateId"]: raise RuntimeError("没有识别到 cateId，请粘贴图1完整网址")
            self.log(f"类目 cateId={cc['cateId']}｜过去30天={self.dr}｜年同比")
            mp=1; rank=0; processed=0
            while not self.stop and mp<=100:
                q={"dateRange":self.dr,"dateType":"day","page":mp,"pageSize":50,"order":"desc","marketVersion":"free","parentCateId":cc["parentCateId"],"cateId":cc["cateId"],"cateFlag":cc["cateFlag"],"kwType":"search","rankType":"hot","orderBy":"seIpvUvHits","device":0}
                rows,total=await self.api(page,RANK,q)
                if not rows: break
                done=False
                for it in rows:
                    if self.stop: break
                    if not isinstance(it,dict): continue
                    rank+=1; kw=compact(metric(it.get("searchWord"))); pop=metric(it.get("seIpvUvHits"))
                    if not kw: continue
                    u=upper(pop)
                    if u is not None and u<self.mn: self.log(f"图1达到阈值：{kw} {pop} < {self.mn}，停止"); done=True; break
                    processed+=1; self.status(f"处理中：{kw}｜主词 {processed}｜排名 {rank}"); self.log(f"主词 #{rank}：{kw}（{pop}）")
                    rr=await self.related(page,kw,mp,rank,pop); store.add(rr); self.log(f"已保存 {len(rr)} 条关联词"); await asyncio.sleep(random.uniform(.7,1.2))
                if done or self.stop or len(rows)<50 or (total and mp*50>=total): break
                mp+=1; await asyncio.sleep(random.uniform(.5,.9))
        self.log(f"完成：{path}"); return str(path)

class App:
    def __init__(self,root):
        self.root=root; root.title(APP); root.geometry("1000x720"); self.worker=None; self.collector=None
        f=ttk.Frame(root,padding=16); f.pack(fill="both",expand=True)
        ttk.Label(f,text=APP,font=("Microsoft YaHei UI",18,"bold")).grid(row=0,column=0,columnspan=4,sticky="w",pady=(0,14))
        ttk.Label(f,text="搜索排行 URL").grid(row=1,column=0,sticky="w"); self.url=tk.StringVar(); ttk.Entry(f,textvariable=self.url,width=92).grid(row=1,column=1,columnspan=3,sticky="ew",pady=5)
        ttk.Label(f,text="图1主词最低搜索人气").grid(row=2,column=0,sticky="w"); self.mn=tk.IntVar(value=300); ttk.Entry(f,textvariable=self.mn,width=12).grid(row=2,column=1,sticky="w")
        ttk.Label(f,text="图2关联词最低搜索人气").grid(row=2,column=2,sticky="w"); self.rn=tk.IntVar(value=100); ttk.Entry(f,textvariable=self.rn,width=12).grid(row=2,column=3,sticky="w")
        ttk.Label(f,text="数据周期").grid(row=3,column=0,sticky="w"); ttk.Label(f,text="过去30个完整自然日").grid(row=3,column=1,sticky="w"); ttk.Label(f,text="对比方式").grid(row=3,column=2,sticky="w"); ttk.Label(f,text="年同比").grid(row=3,column=3,sticky="w")
        ttk.Label(f,text="导出目录").grid(row=4,column=0,sticky="w"); self.od=tk.StringVar(value=str(OUT)); ttk.Entry(f,textvariable=self.od).grid(row=4,column=1,columnspan=2,sticky="ew",pady=5); ttk.Button(f,text="选择目录",command=self.pick).grid(row=4,column=3,sticky="w")
        b=ttk.Frame(f); b.grid(row=5,column=0,columnspan=4,sticky="w",pady=(12,8))
        for t,cmd in [("启动采集专用 Chrome",self.chrome),("测试 Chrome",self.test),("开始采集",self.start),("停止",self.stop),("打开导出目录",self.openout)]: ttk.Button(b,text=t,command=cmd).pack(side="left",padx=4)
        self.st=tk.StringVar(value="API版：浏览器只负责登录，数据直接读取生意参谋接口"); ttk.Label(f,textvariable=self.st).grid(row=6,column=0,columnspan=4,sticky="w")
        ttk.Label(f,text="运行日志").grid(row=7,column=0,columnspan=4,sticky="w",pady=(10,2)); self.box=tk.Text(f,height=25,wrap="word"); self.box.grid(row=8,column=0,columnspan=4,sticky="nsew")
        ttk.Label(f,text="首次：启动专用 Chrome → 正常登录生意参谋 → 粘贴图1完整网址 → 开始采集。",foreground="#555").grid(row=9,column=0,columnspan=4,sticky="w",pady=(10,0)); f.rowconfigure(8,weight=1)
        for i in range(4): f.columnconfigure(i,weight=1)
    def log(self,s): self.root.after(0,lambda x=s:(self.box.insert("end",x+"\n"),self.box.see("end")))
    def status(self,s): self.root.after(0,lambda x=s:self.st.set(x))
    def pick(self):
        p=filedialog.askdirectory(initialdir=self.od.get())
        if p:self.od.set(p)
    def openout(self): Path(self.od.get()).mkdir(parents=True,exist_ok=True); os.startfile(self.od.get())
    def chrome(self):
        cs=[Path(os.environ.get("ProgramFiles",""))/"Google/Chrome/Application/chrome.exe",Path(os.environ.get("ProgramFiles(x86)",""))/"Google/Chrome/Application/chrome.exe",Path(os.environ.get("LocalAppData",""))/"Google/Chrome/Application/chrome.exe"]
        c=next((x for x in cs if x.exists()),None)
        if not c: messagebox.showerror("未找到 Chrome","请先安装 Google Chrome"); return
        prof=ROOT/"chrome-profile"; prof.mkdir(parents=True,exist_ok=True); subprocess.Popen([str(c),"--remote-debugging-port=9222","--remote-allow-origins=*",f"--user-data-dir={prof}","https://sycm.taobao.com/"]); self.st.set("已启动专用 Chrome；首次请手动登录")
    def test(self):
        def w():
            async def go():
                async with async_playwright() as p:
                    b=await p.chromium.connect_over_cdp("http://127.0.0.1:9222"); return [pg.url for c in b.contexts for pg in c.pages]
            try:
                u=asyncio.run(go()); self.log("Chrome 连接成功："+" | ".join(u[:4])); self.status("Chrome 连接成功"); self.root.after(0,lambda:messagebox.showinfo("成功","Chrome 已连接"))
            except Exception as e:
                er=str(e); self.log("Chrome 连接失败："+er); self.root.after(0,lambda x=er:messagebox.showerror("失败","请先启动采集专用 Chrome。\n\n"+x))
        threading.Thread(target=w,daemon=True).start()
    def start(self):
        if self.worker and self.worker.is_alive(): messagebox.showwarning("提示","任务正在运行"); return
        if not compact(self.url.get()): messagebox.showwarning("缺少 URL","请粘贴图1完整网址"); return
        self.collector=Collector(self.url.get(),self.mn.get(),self.rn.get(),self.od.get(),self.log,self.status); self.worker=threading.Thread(target=self._run,daemon=True); self.worker.start()
    def _run(self):
        try:
            p=asyncio.run(self.collector.run()); self.status("完成："+p); self.root.after(0,lambda x=p:messagebox.showinfo("完成","文件已保存：\n"+x))
        except Exception as e:
            er=str(e); self.log("运行失败："+er); self.status("运行失败"); self.root.after(0,lambda x=er:messagebox.showerror("运行失败",x))
    def stop(self):
        if self.collector:self.collector.stop=True; self.status("正在停止…")

def selftest():
    assert upper("300 ~ 600")==600 and upper("1万 ~ 2万")==20000 and upper(123)==123
    assert norm_url("/search_rank?a=1")=="https://sycm.taobao.com/mc/free/search_rank?a=1"
    assert ctx("https://sycm.taobao.com/mc/free/search_rank?cateId=2")["cateId"]=="2"
    assert len(window30())==21
    print("SELF_TEST_OK")

if __name__=="__main__":
    if "--self-test" in sys.argv:selftest()
    else:
        r=tk.Tk(); App(r); r.mainloop()
