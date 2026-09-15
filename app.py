
import asyncio
import os
import re
import subprocess
import threading
import time
import traceback
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from openpyxl import Workbook
from playwright.async_api import async_playwright

APP_NAME = "ç”Ÿæ„å‚è°‹èµ›é“æ•°æ®é‡‡é›†å™¨"
DATA_DIR = Path.home() / ".sycm_collector_v3"
EXPORT_DIR = DATA_DIR / "exports"
LOG_FILE = DATA_DIR / "collector.log"
for p in (DATA_DIR, EXPORT_DIR):
    p.mkdir(parents=True, exist_ok=True)

@dataclass
class Config:
    url: str = ""
    main_min_popularity: int = 300
    related_min_popularity: int = 100
    period: str = "30å¤©"
    compare: str = "å¹´åŒæ¯”"
    output_dir: str = str(EXPORT_DIR)
    cdp_url: str = "http://127.0.0.1:9222"

class Logger:
    def __init__(self, cb=None): self.cb = cb
    def log(self, msg):
        line = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception: pass
        if self.cb: self.cb(line)

def compact(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def pop_range(text):
    s = str(text or "").replace(",", "").strip()
    m = re.search(r'([0-9.]+)\s*(ä¸‡|åƒ)?\s*[~ï½\-â€”åˆ°]\s*([0-9.]+)\s*(ä¸‡|åƒ)?', s)
    def cv(v,u):
        x=float(v)
        if u=="ä¸‡": x*=10000
        elif u=="åƒ": x*=1000
        return x
    if m:
        return cv(m.group(1),m.group(2)), cv(m.group(3),m.group(4))
    m = re.search(r'([0-9.]+)\s*(ä¸‡|åƒ)?', s)
    if m:
        x=cv(m.group(1),m.group(2))
        return x,x
    return None,None

def below_threshold(text, threshold):
    lo, hi = pop_range(text)
    if hi is None: return False
    # ä¿å®ˆè§„åˆ™ï¼šæ•´ä¸ªåŒºé—´éƒ½ä½äºé˜ˆå€¼æ‰åœæ­¢
    return hi < threshold

class ExcelStore:
    headers = [
        "é‡‡é›†æ—¥æœŸ","æ¥æºURL","æ¥æºé¡µç ","æ¥æºæœç´¢è¯","å…³è”æœç´¢è¯",
        "æœç´¢äººæ°”","æœç´¢äººæ°”åŒæ¯”","ç‚¹å‡»ç‡","ç‚¹å‡»ç‡åŒæ¯”",
        "æ”¯ä»˜è½¬åŒ–ç‡","æ”¯ä»˜è½¬åŒ–ç‡åŒæ¯”","æ”¯ä»˜ä¹°å®¶æ•°","æ”¯ä»˜ä¹°å®¶æ•°åŒæ¯”",
        "éœ€æ±‚ä¾›ç»™æ¯”","éœ€æ±‚ä¾›ç»™æ¯”åŒæ¯”","å¤©çŒ«å•†å“ç‚¹å‡»å æ¯”","å¤©çŒ«å•†å“ç‚¹å‡»å æ¯”åŒæ¯”"
    ]
    def __init__(self, path):
        self.path=Path(path)
        self.wb=Workbook()
        self.ws=self.wb.active
        self.ws.title="åŸå§‹æ•°æ®"
        self.ws.append(self.headers)
        self.ws.freeze_panes="A2"
        for i,w in enumerate([13,42,9,24,28,16,16,14,14,16,16,16,16,16,16,22,22],1):
            self.ws.column_dimensions[self.ws.cell(1,i).column_letter].width=w
        self.wb.save(self.path)
    def add(self, rows):
        for r in rows:
            self.ws.append([r.get(h,"") for h in self.headers])
        self.wb.save(self.path)

class Collector:
    def __init__(self,cfg,logger,status_cb=None):
        self.cfg=cfg; self.logger=logger; self.status_cb=status_cb
        self.pause_evt=threading.Event(); self.pause_evt.set()
        self.stop_evt=threading.Event()

    def pause(self): self.pause_evt.clear(); self.logger.log("å·²æš‚åœ")
    def resume(self): self.pause_evt.set(); self.logger.log("ç»§ç»­")
    def stop(self): self.stop_evt.set(); self.pause_evt.set(); self.logger.log("è¯·æ±‚åœæ­¢")

    async def gate(self):
        while not self.pause_evt.is_set():
            await asyncio.sleep(.25)
        if self.stop_evt.is_set():
            raise asyncio.CancelledError()

    async def connect(self,p):
        self.logger.log("æ­£åœ¨è¿æ¥é‡‡é›†ä¸“ç”¨ Chromeâ€¦")
        browser=await p.chromium.connect_over_cdp(self.cfg.cdp_url)
        if not browser.contexts:
            raise RuntimeError("Chrome å·²è¿æ¥ï¼Œä½†æ²¡æœ‰æµè§ˆå™¨ä¸Šä¸‹æ–‡ã€‚")
        ctx=browser.contexts[0]
        pages=ctx.pages
        # ä¼˜å…ˆé€‰æ‹©å·²æ‰“å¼€çš„ç”Ÿæ„å‚è°‹æœç´¢æ’è¡Œé¡µ
        sycm=[pg for pg in pages if "sycm.taobao.com" in pg.url]
        page=sycm[-1] if sycm else (pages[-1] if pages else await ctx.new_page())
        return browser,ctx,page

    async def is_ranking_page(self,page):
        try:
            if await page.get_by_text("æœç´¢æ’è¡Œ",exact=True).count():
                if await page.locator("table tbody tr").count():
                    return True
        except Exception: pass
        return False

    async def ensure_page(self,page):
        if self.cfg.url:
            self.logger.log("æ‰“å¼€æŒ‡å®šæœç´¢æ’è¡Œ URLâ€¦")
            await page.goto(self.cfg.url,wait_until="domcontentloaded",timeout=30000)
            await asyncio.sleep(1)
        for _ in range(60):
            if await self.is_ranking_page(page):
                return
            await asyncio.sleep(1)
        raise RuntimeError("æ²¡æœ‰æ£€æµ‹åˆ°â€œæœç´¢æ’è¡Œâ€è¡¨æ ¼ã€‚è¯·å…ˆåœ¨é‡‡é›†ä¸“ç”¨ Chrome ä¸­æ‰‹åŠ¨è¿›å…¥å›¾1é¡µé¢ã€‚")

    async def get_table_rows(self,page):
        # é¡µé¢å¯èƒ½æœ‰å¤šä¸ªè¡¨æ ¼ï¼Œä¼˜å…ˆå–å«â€œæœç´¢è¯/æœç´¢äººæ°”/æœç´¢åˆ†æâ€çš„è¡¨
        tables=page.locator("table")
        n=await tables.count()
        best=None
        for i in range(n):
            t=tables.nth(i)
            try:
                txt=compact(await t.inner_text())
                score=sum(x in txt for x in ["æœç´¢è¯","æœç´¢äººæ°”","æœç´¢åˆ†æ"])
                if score>=2:
                    best=t; break
            except Exception: pass
        if best is None:
            best=page.locator("table").first
        return best.locator("tbody tr")

    async def ranking_rows(self,page):
        rows=await self.get_table_rows(page)
        n=await rows.count()
        out=[]
        for i in range(n):
            r=rows.nth(i)
            try:
                raw=await r.locator("td").all_inner_texts()
                cells=[compact(x) for x in raw]
                if len(cells)<3: continue
                keyword=""
                # æ’åé€šå¸¸ç¬¬1åˆ—ï¼Œå…³é”®è¯é€šå¸¸ç¬¬2åˆ—ï¼›å…ˆä¼˜å…ˆç¬¬2åˆ—
                for idx in [1,0,2,3]:
                    if idx>=len(cells): continue
                    c=cells[idx]
                    if c and not re.fullmatch(r'[\d\s.%~ï½+\-ä¸‡åƒ]+',c) and c not in ("æœç´¢åˆ†æ","å•†æœºå‘ç°","è¶‹åŠ¿"):
                        keyword=c; break
                pop=""
                for c in cells:
                    if re.search(r'\d',c) and (re.search(r'[~ï½]',c) or "ä¸‡" in c or "åƒ" in c):
                        pop=c; break
                if keyword and pop:
                    out.append({"keyword":keyword,"pop":pop})
            except Exception: pass
        return out

    async def click_analysis(self,page,keyword):
        rows=await self.get_table_rows(page)
        n=await rows.count()
        for i in range(n):
            r=rows.nth(i)
            try:
                if keyword not in compact(await r.inner_text()):
                    continue
                link=r.get_by_text("æœç´¢åˆ†æ",exact=True)
                if not await link.count():
                    continue

                # åŒæ—¶å…¼å®¹åŒé¡µè·³è½¬å’Œæ–°æ ‡ç­¾é¡µ
                old_pages=len(page.context.pages)
                old_url=page.url
                try:
                    async with page.context.expect_page(timeout=1500) as pi:
                        await link.first.click(timeout=5000)
                    newp=await pi.value
                    await newp.wait_for_load_state("domcontentloaded",timeout=10000)
                    return newp, True
                except Exception:
                    await asyncio.sleep(1)
                    if page.url != old_url or await page.get_by_text("ç›¸å…³åˆ†æ" , exact=True).count():
                        return page, False
            except Exception:
                continue
        return None,False

    async def set_filters(self,page):
        # åªåœ¨é¡µé¢ä¸Šæ–¹ç«™å¯èƒ½æœ‰çš„åˆ†æåŒºåŸŸæ“ä½œï¼Œé¿å…ä¹±å‹¾å…¶å®ƒ Checkbox
        self.logger.log(f"è®¾ç½®ï¼šs{lelf.cfg.compare} / {self.cfg.period} / æŒ‡æ ‡å…¨éƒ¨")
        try:
            loc=page.get_by_text(self.cfg.compare,exact=True)
            if await loc.count():
                await loc.last.click(timeout=5000)
                await asyncio.sleep(.5)
        except Exception as e:
            self.logger.log("æé†’ï¼šæœªç¡®è®¤â€œå¹´åŒæ¯”â€æ˜¯å¦å·²åˆ‡æ¢ã€‚")
        try:
            loc=page.get_by_text(self.cfg.period,exact=True)
            if await loc.count():
                await loc.last.click(timeout=5000)
                await asyncio.sleep(.5)
        except Exception:
            self.logger.log("æé†’ï¼šæŠ¬ç« ç†¶ä¸æ‰¾åˆ°â€œ30å¤©â€æŒ‰é’®ã€‚")

        wanted=["æœç´¢äººæ°”","ç‚¹å‡»ç‡","æ”¯ä»˜è½¬åŒ–ç‡","æ”¯ä»˜ä¹°å®¶æ•°","éœ€æ±‚ä¾›ç»™æ¯”","å¤©çŒ«å•†å“ç‚¹å‡»å æ¯”"]
        # å…ˆæ£€æŸ¥æµè§ˆå™¨äº‹æ˜Ÿå¹å¶å…° checkboyï¼›ä¸å†ä¸ªè¡¨æ‰€æœ‰àcheckbox ä¸€è‚é‡£ä¸ŠåŠ¾ä¸Š
        for txt in wanted:
            try:
                lab=page.get_by_text(txt,exact=True)
                if not await lab.count(): continue
                el=lab.first
                parent=el.locator("xpath=..")
                cb=parent.locator('input[type="checkbox"]')
                if await cb.count():
                    c=cb.first
                    if not await c.is_checked():
                        await c.check(force=True)
            except Exception:
                pass
        await asyncio.sleep(1.2)

    def split_metric(self,text):
        # ä¿ç•™æ¢è¡Œï¼›é¢‘åŸºé€šå¸¸äº“ â€œå½“å‰å€¼\nåŒæ¯”â€
        raw=str(text or "").strip()
        parts=[compact(x) for x in re.split(r'[\r\n]+',raw) if compact(x)]
        if len(parts)>=2:
            return parts[0],parts[-1]
        s=compact(raw)
        # ä¾‹å¦‚ "1ä¸‡ ~ 2ä¸‡ 20%" / "92% +9.52%"
        m=re.match(r'(.+?)\s+([+\-]?\d[\s,.]*%)$',s)
        if m:
            return m.group(1).strip(),m.group(2).strip()
        return s,""

    async def analysis_rows(self,page,source_kw,source_url,page_no):
        # æ‰¾å«ä¸ªæŒ‡æ ‡åçš„åˆ†æè¡¡
        tables=page.locator("table")
        tn=await tables.count()
        table=None
        for i in range(tn):
            t=tables.nth(i)
            try:
                txt=compact(await t.inner_text())
                if "æœç´¢äººæ°”" in txt and "æ”¯ä»˜ä¹³å®¶æ•°" in txt and "éœ€æ±‚ä¾›ç»™æ¯”" in txt:
                    table=t; break
            except Exception: pass
        if table is None:
            raise RuntimeError("å›¾2æ²¡æœ‰è¯†åˆ«åˆ°â€œç›¸å…³åˆ†æâ€æ•°æ®è¡¨æ€")

        rows=table.locator("tbody tr")
        n=await rows.count()
        result=[]
        for i in range(n):
            await self.gate()
            r=rows.nth(i)
            try:
                raw=await r.locator("td").all_inner_texts()
                cells=[str(x).strip() for x in raw]
                if len(cells)<7: continue

                # ç¬¬ä¸€åˆ—é€šå¸¸æ˜¯å…³å¸¢è¿Ÿå¥½ä¸ä¸“åŸ¹
                kw=compact(cells[0]).replace("æœç´¢è¯","").strip()
                metrics=[self.split_metric(cells[j]) for j in range(1,min(7,len(cells)))]
                while len(metrics)<6: metrics.append(("",""))

                pop_current,pop_yoy=metrics[0]
                if below_threshold(pop_current,self.cfg.related_min_popularity):
                    self.logger.log(f"å›¾2ï¼š{ZŞßH9¤'9í(¹.®¹¬%ÜÜØİ\œ™[H9mì¹/c¹.£ˆÜÙ[‹˜Ù™Ëœ™[]YÛZ[—ÜÜ[\š]_{ï#9`g9«i¹odùbcz+ãy¢jyìexà ˆŠBˆœ™XZÂ‚ˆ™\İ[˜\[™
Âˆºaáúfá¹¥éy¥¡È[YKœİ™[YJ‰VKI[KIYŠKˆ¹§iy®¤T“œÛİ\˜ÙWİ\›ˆ¹§iy®¤:hmyè HœYÙWÛ›Ëˆ¹§iy®¤9¤'9í(º+ãHœÛİ\˜ÙWÚİËˆ¹alú e9¤'9í(º+ãHšİËˆ¹¤'9í(¹.®¹¬%œÜØİ\œ™[ˆ¹¤'9í(¹.®¹¬%9d#9«åœÜŞ[ŞKˆ¹à®yaîùã¡È›Y]šXÜÖÌWVÌK¹à®yaîùã¡ùd#9«å›Y]šXÜÖÌWVÌWKˆ¹¥+ù.æ:/k9c%¹ã¡È›Y]šXÜÖÌ—VÌK¹¥+ù.æ:/k9c%¹£¡ùd#9«å›Y]šXÜÖÌ—VÌWKˆ¹¥+ù.æ9.l9k­¹¥l›Y]šXÜÖÌ×VÌK¹¥+ù.æ9.l9k­¹¥l9d#9«å›Y]šXÜÖÌ×VÌWKˆºg 9¬`¹/¦ùîæy¥¡È›Y]šXÜÖÍVÌKºg 9¬`¹/¦ùîæy«å9d#9«å›Y]šXÜÖÍVÌWKˆ¹i*yã*ùea¹dàyà®yaîùch9«å›Y]šXÜÖÍWVÌK¹i*yã*ùea¹dàyà®yaîùchy«å9d#9«å›Y]šXÜÖÍWVÌWKˆJBˆ^Ù\^Ù\[Û‚ˆÛÛ[YBˆ™]\›ˆ™\İ[‚ˆ\Ş[˜ÈYˆ™^ÜYÙJÙ[‹YÙJN‚ˆÈ9ao9k®x '9."ù. :hmx 'y¥¡Ékeù¢%ˆ]KØ\šXK[X™[ˆØ[™Y]\ÏVÂˆYÙK™Ù]ØWİ^
¹."ù. :hmH‹^XİUYJKˆYÙK›ØØ]ÜŠ	Öİ]OH¹."ù. :hmH—IÊKˆYÙK›ØØ]ÜŠ	ÖØ\šXK[X™[H¹."ù. :hmH—IÊBˆBˆ›ÜˆØÈ[ˆØ[™Y]\Î‚ˆN‚ˆYˆ›İ]ØZ]ØË˜Ûİ[

NˆÛÛ[YBˆ[[ØË›\İˆÛÏJ]ØZ][™Ù]Ø]šX]J˜Û\ÜÈŠHÜˆˆŠK›İÙ\Š
Bˆ\ØX›YJ]ØZ][™Ù]Ø]šX]J™\ØX›YŠJH\È›İ›Û™Bˆ\šXOJ]ØZ][™Ù]Ø]šX]J˜\šXKY\ØX›YŠHÜˆˆŠK›İÙ\Š
OOHYH‚ˆYˆ\ØX›YÜˆ\šXHÜˆ™\ØX›Yˆ[ˆÛÎ‚ˆ™]\›ˆ˜[ÙBˆ]ØZ][˜ÛXÚÊ[Y[İ]ML
Bˆ]ØZ]\Ş[˜Ú[ËœÛY\
JBˆ™]\›ˆYBˆ^Ù\^Ù\[Ûˆ\ÜÂˆ™]\›ˆ˜[ÙB‚ˆ\Ş[˜ÈYˆ[ŠÙ[ŠN‚ˆİ]\T]
Ù[‹˜Ù™Ë›İ]]Ù\ŠNÈİ]\‹›ZÙ\Š\™[ÏUYK^\İÛÚÏUYJBˆİ]š[O[İ]\‹Ùˆ¹å'ù¡#ùcàº,"×ù¤'9í(¹b!¹§¤Şİ[YKœİ™[YJ	ÉVI[IYÉRIS')}.xlsx"
        store=ExcelStore(outfile)

        async with async_playwright() as p:
            browser,ctx,page=await self.connect(p)
            await self.ensure_page(page)
            source_url=page.url
            self.logger.log("Chrome è¿æ¥æˆåŠŸï¼Œæ£€æµ‹åˆ°æœç´¢æ’è¡Œã€‚")
            page_no=1
            processed=0
            global_stop=False

            while not global_stop:
                await self.gate()
                items=await self.ranking_rows(page)
                if not items:
                    raise RuntimeError(f"ç¬¬ {page_no} é¡µæ²¡æœ‰è¯†åˆ«åˆ°ä¸»é…³é”®è¯ã€‚")

                for item in items:
                    await self.gate()
                    kw=item["keyword"]; pop=item["pop"]
                    if below_threshold(pop,self.cfg.main_min_popularity):
                        self.logger.log(f"å›¾1ï¼š{ZŞßH9¤'9í(¹.®¹¬%ÜÜH9mì¹/c¹.£ˆÜÙ[‹˜Ù™Ë›XZ[—ÛZ[—ÜÜ[\š]_{ï#9`g9«h¹d#¹îëy..ú+ãxà ˆŠBˆÛØ˜[ÜİÜUYBˆœ™XZÂ‚ˆ›ØÙ\ÜÙY
ÏLBˆYˆÙ[‹œİ]\×ØØœÙ[‹œİ]\×ØØŠˆ¹i!9ä!¹.+ûï&Ö··ÒûÙÎ[{.ZHNyb·&ö6W76VGÒKŠ®K‹¾ŠøÒ"¢6VÆbæÆövvW"æÆör†b.Y»ãK‹¾ŠøŞûÉ§¶·wÒ‡·÷Ò’" ¢æÇ—6—5÷vRÆ÷VæVEöæWsÖv—B6VÆbæ6Æ–6µöæÇ—6—2‡vRÆ·r¢–bæÇ—6—5÷vR—2æöæS ¢6VÆbæÆövvW"æÆör†b.iÊ®ˆ;Ş‹ù¾XZ^8Ç¶·wÎ8Şi	Î{J.XˆniéûÈÎ‹{>‹ørâ"¢6öçF–çVP ¢v—B6VÆbç6WEöf–ÇFW'2†æÇ—6—5÷vR¢&÷w3Öv—B6VÆbææÇ—6—5÷&÷w2†æÇ—6—5÷vRÆ·rÇ6÷W&6U÷W&ÂÇvUöæò¢7F÷&RæFB‡&÷w2¢6VÆbæÆövvW"æÆör†b.[{.KùŞZÙ8Ç¶·wÎ8ŞX[>[‰®‹ùşŠøÒ¶ÆVâ‡&÷w2—ÒiÚ8"" ¢2Y¹îX‹hé.ŠÎšP¢–b÷VæVEöæWs ¢G'“¢v—BæÇ—6—5÷vRæ6Æ÷6R‚¢W†6WBW†6WF–öã¢70¢VÇ6S ¢G'“ ¢v—BvRævõö&6²‡v—E÷VçF–ÃÒ&FöÖ6öçFVçFÆöFVB"ÇF–ÖV÷WCÓ#¢v—B7–æ6–òç6ÆVWƒ¢W†6WBW†6WF–öã ¢v—BvRæv÷Fò‡6÷W&6U÷W&ÂÇv—E÷VçF–ÃÒ&FöÖ6öçFVçFÆöFVB"ÇF–ÖV÷WCÓ#¢v—B7–æ6–òç6ÆVWƒ ¢2zîKùŞ‹ùNY¹îjÚ>zîš^ûÉ¾ˆº^XènXû.‹ùNY¹îZK‹J^ûÈÎy»Nhê^Y¹îk©U$À¢–bæ÷Bv—B6VÆbæ—5÷&æ¶–æu÷vR‡vR“ ¢v—BvRæv÷Fò‡6÷W&6U÷W&ÂÇv—E÷VçF–ÃÒ&FöÖ6öçFVçFÆöFVB"ÇF–ÖV÷WCÓ#¢v—B7–æ6–òç6ÆVWƒ ¢–bvÆö&Å÷7F÷¢'&V°¢–bæ÷Bv—B6VÆbææW‡E÷vR‡vR“ ¢6VÆbæÆövvW"æÆör‚.[{.X‹iÈYîKˆš^8""¢'&V°¢vUöæò³Ó ¢6VÆbæÆövvW"æÆör†b.˜x~™¸nZèÎh‰ûÉ§¶÷WFf–ÆWÒ"¢&WGW&â7G"†÷WFf–ÆR ¦6Æ72 ¢FVbõö–æ—Eõò‡6VÆbÇ&ö÷B“ ¢6VÆbç&ö÷C×&ö÷@¢&ö÷BçF—FÆR„ôäÔR¢&ö÷BævVöÖWG'’‚#ƒs3"¢6VÆbæ6öÆÆV7F÷#ÔæöæS²6VÆbçv÷&¶W#ÔæöæP ¢c×GF²äg&ÖR‡&ö÷BÇFF–æsÓb“²bç6²†f–ÆÃÒ&&÷F‚"ÆW‡æCÕG'VR¢GF²äÆ&VÂ†bÇFW‡CÔôäÔRÆföçCÒ‚$Ö–7&÷6ögB–†V’T’"Ã‚Â&&öÆB"’æw&–B‡&÷sÓÆ6öÇVÖãÓÆ6öÇVÖç7ãÓBÇ7F–6·“Ò'r"ÇG“ÒƒÃB’ ¢GF²äÆ&VÂ†bÇFW‡CÒ.i	Î{J.hé.ŠÂU$ÎûÈ„6‡&öÖR[{.XÎYÊY»ãi{nXúşyYz›®ûÈ’"’æw&–B‡&÷sÓÆ6öÇVÖãÓÇ7F–6·“Ò'r"¢6VÆbçW&Ã×F²å7G&–æuf"‚¢GF²äVçG'’†bÇFW‡Gf&–&ÆS×6VÆbçW&ÂÇv–GFƒÓ“’æw&–B‡&÷sÓÆ6öÇVÖãÓÆ6öÇVÖç7ãÓ2Ç7F–6·“Ò&Wr"ÇG“ÓR ¢GF²äÆ&VÂ†bÇFW‡CÒ.Y»ãK‹¾ŠøŞiÈKØîi	Î{J.K«®k	B"’æw&–B‡&÷sÓ"Æ6öÇVÖãÓÇ7F–6·“Ò'r"¢6VÆbæÖ–ã×F²ä–çEf"‡fÇVSÓ3“²GF²äVçG'’†bÇFW‡Gf&–&ÆS×6VÆbæÖ–âÇv–GFƒÓ"’æw&–B‡&÷sÓ"Æ6öÇVÖãÓÇ7F–6·“Ò'r" ¢GF²äÆ&VÂ†bÇFW‡CÒ.Y»ã.X[>ˆNŠøŞiÈKØîi	Î{J.K«®k	B"’æw&–B‡&÷sÓ"Æ6öÇVÖãÓ"Ç7F–6·“Ò'r"¢6VÆbç&VÃ×F²ä–çEf"‡fÇVSÓ“²GF²äVçG'’†bÇFW‡Gf&–&ÆS×6VÆbç&VÂÇv–GFƒÓ"’æw&–B‡&÷sÓ"Æ6öÇVÖãÓ2Ç7F–6·“Ò'r" ¢GF²äÆ&VÂ†bÇFW‡CÒ.YiÉò"’æw&–B‡&÷sÓ2Æ6öÇVÖãÓÇ7F–6·“Ò'r"¢6VÆbçW&–öC×F²å7G&–æuf"‡fÇVSÒ#3ZJ’"¢GF²ä6öÖ&ö&÷‚†bÇFW‡Gf&–&ÆS×6VÆbçW&–öBÇfÇVW3Õ²#~ZJ’"Â#3ZJ’%ÒÇ7FFSÒ'&VFöæÇ’"Çv–GFƒÓ’æw&–B‡&÷sÓ2Æ6öÇVÖãÓÇ7F–6·“Ò'r" ¢GF²äÆ&VÂ†bÇFW‡CÒ.ZûjùB"’æw&–B‡&÷sÓ2Æ6öÇVÖãÓ"Ç7F–6·“Ò'r"¢6VÆbæ6ö×&S×F²å7G&–æuf"‡fÇVSÒ.[›NYÎjùB"¢GF²ä6öÖ&ö&÷‚†bÇFW‡Gf&–&ÆS×6VÆbæ6ö×&RÇfÇVW3Õ².[›NYÎjùB"Â.xêşjùB%ÒÇ7FFSÒ'&VFöæÇ’"Çv–GFƒÓ’æw&–B‡&÷sÓ2Æ6öÇVÖãÓ2Ç7F–6·“Ò'r" ¢GF²äÆ&VÂ†bÇFW‡CÒ.ZûÎX{®yºî[ÙR"’æw&–B‡&÷sÓBÆ6öÇVÖãÓÇ7F–6·“Ò'r"¢6VÆbæ÷WC×F²å7G&–æuf"‡fÇVS×7G"„U…õ%EôD•"’¢GF²äVçG'’†bÇFW‡Gf&–&ÆS×6VÆbæ÷WBÇv–GFƒÓc’æw&–B‡&÷sÓBÆ6öÇVÖãÓÆ6öÇVÖç7ãÓ"Ç7F–6·“Ò&Wr"ÇG“ÓR¢GF²ä'WGFöâ†bÇFW‡CÒ.˜hºyºî[ÙR"Æ6öÖÖæC×6VÆbç–6²’æw&–B‡&÷sÓBÆ6öÇVÖãÓ2Ç7F–6·“Ò'r" ¢#×GF²äg&ÖR†b“²"æw&–B‡&÷sÓRÆ6öÇVÖãÓÆ6öÇVÖç7ãÓBÇ7F–6·“Ò'r"ÇG“Òƒ"Ã‚’¢GF²ä'WGFöâ†"ÇFW‡CÒ.Y
şXª˜x~™¸nK‰>yJ‚6‡&öÖR"Æ6öÖÖæC×6VÆbæÆVæ6…ö6‡&öÖR’ç6²‡6–FSÒ&ÆVgB"ÇGƒÒƒÃ‚’¢GF²ä'WGFöâ†"ÇFW‡CÒ.kX¾ŠùR6‡&öÖR‹ùîhêR"Æ6öÖÖæC×6VÆbçFW7Eö6öæâ’ç6²‡6–FSÒ&ÆVgB"ÇGƒÓB¢GF²ä'WGFöâ†"ÇFW‡CÒ.[ÈZx¾˜x~™¸b"Æ6öÖÖæC×6VÆbç7F'B’ç6²‡6–FSÒ&ÆVgB"ÇGƒÓB¢GF²ä'WGFöâ†"ÇFW‡CÒ.i¨.XÂ"Æ6öÖÖæCÖÆÖ&F§6VÆbæ6öÆÆV7F÷"æB6VÆbæ6öÆÆV7F÷"çW6R‚’’ç6²‡6–FSÒ&ÆVgB"ÇGƒÓB¢GF²ä'WGFöâ†"ÇFW‡CÒ.{º~{ºÒ"Æ6öÖÖæCÖÆÖ&F§6VÆbæ6öÆÆV7F÷"æB6VÆbæ6öÆÆV7F÷"ç&W7VÖR‚’’ç6²‡6–FSÒ&ÆVgB"ÇGƒÓB¢GF²ä'WGFöâ†"ÇFW‡CÒ.XÎjÚ""Æ6öÖÖæCÖÆÖ&F§6VÆbæ6öÆÆV7F÷"æB6VÆbæ6öÆÆV7F÷"ç7F÷‚’’ç6²‡6–FSÒ&ÆVgB"ÇGƒÓB¢GF²ä'WGFöâ†"ÇFW‡CÒ.h™>[ÈZûÎX{®yºî[ÙR"Æ6öÖÖæC×6VÆbæ÷Våö÷WB’ç6²‡6–FSÒ&ÆVgB"ÇGƒÓ ¢6VÆbç7FGW3×F²å7G&–æuf"‡fÇVSÒ.XXY
şXª(	Î˜x~™¸nK‰>yJ‚6‡&öÖ^(	ŞûÈÎzÊÎKˆjÊYÊX[nKŠŞh˜¾Xªy›¾[Ù^KˆjÊ8""¢GF²äÆ&VÂ†bÇFW‡Gf&–&ÆS×6VÆbç7FGW2’æw&–B‡&÷sÓbÆ6öÇVÖãÓÆ6öÇVÖç7ãÓBÇ7F–6·“Ò'r"¢6VÆbç#×GF²å&öw&W76&"†bÆÖöFSÒ&–æFWFW&Ö–æFR"“²6VÆbç"æw&–B‡&÷sÓrÆ6öÇVÖãÓÆ6öÇVÖç7ãÓBÇ7F–6·“Ò&Wr"ÇG“ÒƒRÃ’ ¢GF²äÆ&VÂ†bÇFW‡CÒ.‹ùŠÎiz^[ùr"’æw&–B‡&÷sÓ‚Æ6öÇVÖãÓÆ6öÇVÖç7ãÓBÇ7F–6·“Ò'r"¢6VÆbæÆöv&÷ƒ×F²åFW‡B†bÆ†V–v‡CÓ#BÇw&Ò'v÷&B"“²6VÆbæÆöv&÷‚æw&–B‡&÷sÓ’Æ6öÇVÖãÓÆ6öÇVÖç7ãÓBÇ7F–6·“Ò&ç6Wr" ¢æ÷FSÒ‚%c2KˆŞXhŞ[	ŞŠù^Šû¾XùnKÚXéşiÚ^y¨B6‡&öÖR&öf–Æ^8.Zè>KÛşyJKˆKŠ®xºÎz¸¾y¨N(	Î˜x~™¸nK‰>yJ‚6‡&öÖ^(	Ş‹XNiiyºî[Ù^ûÉ¢ ¢.zÊÎKˆjÊKÚh˜¾Xªy›¾[Ù^yIşhHşXø.‹¾ûÈÎK˜¾Yîy›¾[Ù^x«nhKÉ®Kˆy»NKùŞyY8.‹ùj~iz.˜şXXŞˆz®XªXÉnkXşŠxYšy›¾[Ù^š8îhê~ûÈÂ ¢.K™ş˜ş[Èikx˜‚6‡&öÖRZû›¹ŠêNyJh‹~yºî[Ù^‹ùÎzˆ¾‹>Šù^y¨N™™X‹n8""¢GF²äÆ&VÂ†bÇFW‡CÖæ÷FRÇw&ÆVæwFƒÓ“#Æf÷&Vw&÷VæCÒ"3SSR"’æw&–B‡&÷sÓÆ6öÇVÖãÓÆ6öÇVÖç7ãÓBÇ7F–6·“Ò'r"ÇG“ÒƒÃ’ ¢bç&÷v6öæf–wW&Rƒ’ÇvV–v‡CÓ¢f÷"’–â&ævRƒB“¦bæ6öÇVÖæ6öæf–wW&R†’ÇvV–v‡CÓ ¢FVb–6²‡6VÆb“ ¢Öf–ÆVF–Æöræ6¶F—&V7F÷'’†–æ—F–ÆF—#×6VÆbæ÷WBævWB‚’¢–b§6VÆbæ÷WBç6WB‡¢FVb÷Våö÷WB‡6VÆb“ ¢F‚‡6VÆbæ÷WBævWB‚’’æÖ¶F—"‡&VçG3ÕG'VRÆW†—7Eöö³ÕG'VR¢G'“¦÷2ç7F'Ff–ÆR‡6VÆbæ÷WBævWB‚’¢W†6WBW†6WF–öã§70¢FVbÆör‡6VÆbÇ2“§6VÆbç&ö÷BægFW"ƒÆÆÖ&F§6VÆbåöÆör‡2’¢FVböÆör‡6VÆbÇ2“§6VÆbæÆöv&÷‚æ–ç6W'B‚&VæB"Ç2²%Æâ"“·6VÆbæÆöv&÷‚ç6VR‚&VæB"¢FVb7FGW5÷6WB‡6VÆbÇ2“§6VÆbç&ö÷BægFW"ƒÆÆÖ&F§6VÆbç7FGW2ç6WB‡2’  ¢FVbÆVæ6…ö6‡&öÖR‡6VÆb“ ¢6æF–FFW2Ò°¢÷2çF‚æ¦ö–â†÷2æVçf—&öâævWB‚%&öw&Ôf–ÆW2"Â""’Â$vöövÆR"Â$6‡&öÖR"Â$Æ–6F–öâ"Â&6‡&öÖRæW†R"’À¢÷2çF‚æ¦ö–â†÷2æVçf—&öâævWB‚%&öw&Ôf–ÆW2‡ƒƒb’"Â""’Â$vöövÆR"Â$6‡&öÖR"Â$Æ–6F–öâ"Â&6‡&öÖRæW†R"’À¢÷2çF‚æ¦ö–â†÷2æVçf—&öâævWB‚$Æö6ÄFF"Â""’Â$vöövÆR"Â$6‡&öÖR"Â$Æ–6F–öâ"Â&6‡&öÖRæW†R"’À¢Ğ¢6‡&öÖRÒæW‡B‚‡‚f÷"‚–â6æF–FFW2–b‚æB÷2çF‚æW†—7G2‡‚’’ÂæöæR¢–bæ÷B6‡&öÖS ¢ÖW76vV&÷‚ç6†÷vW'&÷"‚.iÊ®h›îX‹6‡&öÖR"Â.k*iÈh›îX‹vöövÆR6‡&öÖ^ûÈÎŠû~XXZèŠ8R6‡&öÖ^8""¢&WGW&à¢&öf–ÆRÒ÷2çF‚æ¦ö–â†÷2æVçf—&öâævWB‚$Æö6ÄFF"Â7G"…F‚æ†öÖR‚’’’Â%5”4Ô6öÆÆV7F÷$6‡&öÖR"¢÷2æÖ¶VF—'2‡&öf–ÆRÂW†—7Eöö³ÕG'VR¢G'“ ¢7V'&ö6W72å÷Vâ…°¢6‡&öÖRÀ¢"Ò×&VÖ÷FRÖFV'Vvv–ær×÷'CÓ“##""À¢"Ò×&VÖ÷FRÖÆÆ÷rÖ÷&–v–ç3Ò¢"À¢b"Ò×W6W"ÖFFÖF—#×·&öf–ÆWÒ"À¢&‡GG3¢ò÷7–6ÒçFö&òæ6öÒò"À¢Ò¢6VÆbç7FGW5÷6WB‚.[{.Y
şXª˜x~™¸nK‰>yJ‚6‡&öÖ^ûÉ¾šinjÊKÛşyJŠû~h˜¾Xªy›¾[Ù^yIşhHşXø.‹²"¢6VÆbæÆör‚.[{.Y
şXª˜x~™¸nK‰>yJ‚6‡&öÖ^ûÈÎ‹>Šù^zºşXú2“##.8""¢W†6WBW†6WF–öâ2S ¢ÖW76vV&÷‚ç6†÷vW'&÷"‚.Y
şXªZK‹JR"Â7G"†R’ ¢FVbFW7Eö6öæâ‡6VÆb“ ¢FVbr‚“ ¢7–æ2FVbB‚“ ¢7–æ2v—F‚7–æ5÷Æ—w&–v‡B‚’2 ¢#Öv—Bæ6‡&öÖ—VÒæ6öææV7Eö÷fW%ö6G‚&‡GG¢òó#rããã£“##""¢W&Ç3Õ·rçW&Âf÷"2–â"æ6öçFW‡G2f÷"r–â2çvW5Ğ¢&WGW&âW&Ç0¢G'“ ¢W&Ç3Ö7–æ6–òç'Vâ‡B‚’¢6VÆbæÆör‚$6‡&öÖR‹ùîhê^h‰X©şûÉ¢"²"Â"æ¦ö–â‡W&Ç5³£UÒ’¢6VÆbç7FGW5÷6WB‚$6‡&öÖR‹ùîhê^h‰X©ò"¢6VÆbç&ö÷BægFW"ƒÆÆÖ&F¦ÖW76vV&÷‚ç6†÷v–æfò‚.h‰X©ò"Â.[{.‹ùîhê^˜x~™¸nK‰>yJ‚6‡&öÖ^8""’¢W†6WBW†6WF–öâ2S ¢W'"Ò7G"†R¢6VÆbæÆör‚$6‡&öÖR‹ùîhê^ZK‹J^ûÉ¢"¶W'"¢6VÆbç7FGW5÷6WB‚$6‡&öÖR‹ùîhê^ZK‹JR"¢6VÆbç&ö÷BægFW"€¢À¢ÆÖ&F×6sÖW'#¢ÖW76vV&÷‚ç6†÷vW'&÷"€¢.‹ùîhê^ZK‹JR"À¢.Šû~XXY
şXª˜x~™¸nK‰>yJ‚6‡&öÖ^ûÈÎ[›nzîŠêB“##"zºşXú>[{.[ÈY
ş8%ÆåÆâ"²×6p¢¢¢F‡&VF–æråF‡&VB‡F&vWC×rÆFVÖöãÕG'VR’ç7F'B‚ ¢FVb7F'B‡6VÆb“ ¢–b6VÆbçv÷&¶W"æB6VÆbçv÷&¶W"æ—5öÆ—fR‚“ ¢ÖW76vV&÷‚ç6†÷wv&æ–ær‚.hùzK¢"Â.K»¾XªjÚ>YÊ‹ùŠÎ8""“·&WGW&à¢6fsÔ6öæf–r€¢W&Ã×6VÆbçW&ÂævWB‚’ç7G&—‚’À¢Ö–åöÖ–å÷÷VÆ&—G“Ö–çB‡6VÆbæÖ–âævWB‚’’À¢&VÆFVEöÖ–å÷÷VÆ&—G“Ö–çB‡6VÆbç&VÂævWB‚’’À¢W&–öC×6VÆbçW&–öBævWB‚’À¢6ö×&S×6VÆbæ6ö×&RævWB‚’À¢÷WGWEöF—#×6VÆbæ÷WBævWB‚’ç7G&—‚’÷"7G"„U…õ%EôD•"¢¢6VÆbæ6öÆÆV7F÷#Ô6öÆÆV7F÷"†6frÄÆövvW"‡6VÆbæÆör’Ç6VÆbç7FGW5÷6WB¢6VÆbç"ç7F'Bƒ¢6VÆbçv÷&¶W#×F‡&VF–æråF‡&VB‡F&vWC×6VÆbå÷'VâÆ&w3Ò†6frÂ’ÆFVÖöãÕG'VR“·6VÆbçv÷&¶W"ç7F'B‚ ¢FVb÷'Vâ‡6VÆbÆ6fr“ ¢G'“ ¢FƒÖ7–æ6–òç'Vâ‡6VÆbæ6öÆÆV7F÷"ç'Vâ‚’¢6VÆbç7FGW5÷6WB‚.ZèÎh‰ûÉ¢"·F‚¢6VÆbç&ö÷BægFW"ƒÆÆÖ&F¦ÖW76vV&÷‚ç6†÷v–æfò‚.ZèÎh‰"Â.˜x~™¸nZèÎh‰ûÉ¥Æâ"·F‚’¢W†6WB7–æ6–òä6æ6VÆÆVDW'&÷# ¢6VÆbç7FGW5÷6WB‚.[{.XÎjÚ""¢W†6WBW†6WF–öâ2S ¢W'"Ò7G"†R¢ÆövvW"‡6VÆbæÆör’æÆör‡G&6V&6²æf÷&ÖEöW†2‚’¢6VÆbç7FGW5÷6WB‚.‹ùŠÎZK‹JR"¢6VÆbç&ö÷BægFW"€¢À¢ÆÖ&F×6sÖW'#¢ÖW76vV&÷‚ç6†÷vW'&÷"‚.‹ùŠÎZK‹JR"Â×6r¢¢f–æÆÇ“ ¢6VÆbç&ö÷BægFW"ƒÇ6VÆbç"ç7F÷ ¦–bõöæÖUõóÓÒ%õöÖ–åõò# ¢&ö÷C×F²åF²‚¢‡&ö÷B¢&ö÷BæÖ–æÆö÷‚ 