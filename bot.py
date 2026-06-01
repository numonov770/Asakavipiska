import os
import re
import logging
import tempfile
import subprocess
import pandas as pd
from html.parser import HTMLParser
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, CommandHandler, filters, ContextTypes

TOKEN = os.environ.get("BOT_TOKEN", "YOUR_TOKEN_HERE")
logging.basicConfig(level=logging.INFO)

def extract_num(s):
    s = str(s).replace(" ", "").replace("\xa0","").replace(",",".")
    m = re.search(r"[\d]+\.?\d*", s)
    return float(m.group()) if m else 0

def clean_name(s):
    parts = str(s).split("/")
    name = parts[2].strip() if len(parts) >= 3 else str(s).strip()
    name = re.sub(r"^\d{5}\s*", "", name).strip()
    name = re.sub(r'^(00634|00668|00508|00599|1900|0200)\s*', "", name).strip()
    return (name[:50] + "...") if len(name) > 50 else (name if name else "Noma'lum")

class TableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.tables = []
        self.cur_table = []
        self.cur_row = []
        self.cur_cell = ""
        self.in_cell = False
    def handle_starttag(self, tag, attrs):
        if tag == "table": self.cur_table = []
        elif tag == "tr": self.cur_row = []
        elif tag in ("td","th"): self.in_cell = True; self.cur_cell = ""
    def handle_endtag(self, tag):
        if tag == "table": self.tables.append(self.cur_table)
        elif tag == "tr":
            if self.cur_row: self.cur_table.append(self.cur_row)
        elif tag in ("td","th"): self.cur_row.append(self.cur_cell.strip()); self.in_cell = False
    def handle_data(self, data):
        if self.in_cell: self.cur_cell += data
    def handle_entityref(self, name):
        if self.in_cell and name == "nbsp": self.cur_cell += ""

def parse_html_file(filepath):
    with open(filepath, "rb") as f:
        raw = f.read()
    content = raw.decode("windows-1251", errors="replace")
    p = TableParser()
    p.feed(content)

    info = {}
    txs = []
    total_debit = 0
    total_credit = 0

    if len(p.tables) >= 1:
        t0 = p.tables[0]
        if len(t0) > 0 and t0[0]: info["bank"] = t0[0][0].strip()
        if len(t0) > 1 and t0[1]: info["period"] = t0[1][0].strip()
        if len(t0) > 2 and t0[2]:
            row2 = t0[2][0].replace("\xa0", " ")
            parts = row2.split("  ")
            info["account"] = parts[0].strip() if parts else row2.strip()
            info["company"] = parts[1].strip() if len(parts) > 1 else ""
            info["inn"] = parts[2].strip() if len(parts) > 2 else ""
        if len(t0) > 3 and t0[3]:
            bal_row = t0[3]
            info["bal_start"] = extract_num(bal_row[0].split(":")[-1]) if bal_row else 0
            info["bal_end"] = extract_num(bal_row[1].split(":")[-1]) if len(bal_row) > 1 else 0

    if len(p.tables) >= 2:
        for row in p.tables[1]:
            if not row or len(row) < 7: continue
            date_str = str(row[0]).strip()
            if "\u0418\u0442\u043e\u0433" in date_str:
                td = extract_num(row[5]) if str(row[5]).strip() else 0
                tc = extract_num(row[6]) if str(row[6]).strip() else 0
                if td > 0: total_debit = td
                if tc > 0: total_credit = tc
                continue
            m = re.match(r"(\d{2}\.\d{2}\.\d{4})", date_str)
            if not m: continue
            date = m.group(1)
            time_m = re.search(r"(\d{2}:\d{2})", date_str)
            time = time_m.group(1) if time_m else ""
            name = clean_name(row[1])
            doc = str(row[2]).strip()
            debit_s = str(row[5]).strip().replace("\xa0","").replace(" ","")
            credit_s = str(row[6]).strip().replace("\xa0","").replace(" ","")
            debit = extract_num(debit_s) if debit_s else None
            credit = extract_num(credit_s) if credit_s else None
            if debit == 0: debit = None
            if credit == 0: credit = None
            desc = str(row[7]).strip() if len(row) > 7 else ""
            desc = re.sub(r"^(00634|00668|00508|00599|1900|0200)\s*", "", desc).strip()
            txs.append({"date": date, "time": time, "doc": doc, "counterparty": name,
                        "debit": debit, "credit": credit, "desc": desc[:100]})

    if total_debit == 0:
        total_debit = sum(tx["debit"] for tx in txs if tx["debit"])
    if total_credit == 0:
        total_credit = sum(tx["credit"] for tx in txs if tx["credit"])
    return info, txs, total_debit, total_credit

def parse_xlsb_xlsx(filepath):
    out_dir = tempfile.mkdtemp()
    subprocess.run(
        ["libreoffice","--headless","--convert-to","xlsx","--outdir",out_dir,filepath],
        capture_output=True, text=True, timeout=30
    )
    xlsx_files = [f for f in os.listdir(out_dir) if f.endswith(".xlsx")]
    if not xlsx_files: raise Exception("Fayl o'qilmadi")
    df = pd.read_excel(os.path.join(out_dir, xlsx_files[0]), sheet_name=0, header=None)

    info = {}
    try:
        info["bank"] = str(df.iloc[0,0]).strip()
        info["period"] = str(df.iloc[1,0]).strip()
        info["account"] = str(df.iloc[2,0]).replace("\xa0"," ").strip()
        info["bal_start"] = extract_num(str(df.iloc[3,0]))
        info["bal_end"] = extract_num(str(df.iloc[3,1]))
        info["company"] = str(df.iloc[3,7]).strip() if pd.notna(df.iloc[3,7]) else ""
    except: pass

    txs = []
    total_debit = 0
    total_credit = 0
    header_row = None
    for i, r in df.iterrows():
        if str(r.iloc[0]).strip() in ("\u0414\u0430\u0442\u0430","Dата","Data"):
            header_row = i
            break
    if header_row is not None:
        for i in range(header_row+1, len(df)):
            r = df.iloc[i]
            date_val = r.iloc[0]
            if pd.isna(date_val): continue
            if "\u0418\u0442\u043e\u0433" in str(date_val):
                try:
                    total_debit = float(r.iloc[5]) if pd.notna(r.iloc[5]) else 0
                    total_credit = float(r.iloc[6]) if pd.notna(r.iloc[6]) else 0
                except: pass
                continue
            try:
                date_fmt = date_val.strftime("%d.%m.%Y") if hasattr(date_val,"strftime") else str(date_val)[:10]
                time_fmt = date_val.strftime("%H:%M") if hasattr(date_val,"strftime") else ""
                doc = str(int(r.iloc[2])) if pd.notna(r.iloc[2]) else ""
                name = clean_name(str(r.iloc[1]) if pd.notna(r.iloc[1]) else "")
                debit = float(r.iloc[5]) if pd.notna(r.iloc[5]) else None
                credit = float(r.iloc[6]) if pd.notna(r.iloc[6]) else None
                desc = str(r.iloc[7]).strip() if pd.notna(r.iloc[7]) else ""
                txs.append({"date":date_fmt,"time":time_fmt,"doc":doc,"counterparty":name,
                            "debit":debit,"credit":credit,"desc":desc[:100]})
            except: continue
    if total_debit == 0: total_debit = sum(tx["debit"] for tx in txs if tx["debit"])
    if total_credit == 0: total_credit = sum(tx["credit"] for tx in txs if tx["credit"])
    return info, txs, total_debit, total_credit

def parse_file(filepath, fname):
    with open(filepath,"rb") as f:
        header = f.read(200)
    is_html = b"<html" in header.lower() or b"<!doc" in header.lower()
    if is_html or fname.lower().endswith(".xls"):
        with open(filepath,"rb") as f:
            raw = f.read()
        if b"<html" in raw[:200].lower() or b"<table" in raw[:500].lower():
            return parse_html_file(filepath)
    return parse_xlsb_xlsx(filepath)

def build_summary(txs):
    debit_by = defaultdict(lambda: {"sum":0,"count":0})
    credit_by = defaultdict(lambda: {"sum":0,"count":0})
    for tx in txs:
        n = tx["counterparty"]
        if tx["debit"]: debit_by[n]["sum"] += tx["debit"]; debit_by[n]["count"] += 1
        if tx["credit"]: credit_by[n]["sum"] += tx["credit"]; credit_by[n]["count"] += 1
    return debit_by, credit_by

def sty(cell, bold=False, bg=None, color="000000", size=10, align="left"):
    cell.font = Font(bold=bold, color=color, size=size, name="Arial")
    if bg: cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    s = Side(style="thin", color="DDDDDD")
    cell.border = Border(left=s, right=s, top=s, bottom=s)

def create_excel(info, txs, total_debit, total_credit, output_path):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Kochirma"
    DARK="1B3A6B"; LBLUE="D6E4F7"; RED="C0392B"; GREEN="1E8449"; WHITE="FFFFFF"; YELLOW="FFF3CD"

    bank = info.get("bank","Bank")
    period = info.get("period","")
    account = info.get("account","")
    company = info.get("company","")
    bal_start = info.get("bal_start",0)
    bal_end = info.get("bal_end",0)

    row = 1
    ws1.merge_cells("A%d:H%d"%(row,row))
    c = ws1.cell(row=row,column=1,value=bank)
    sty(c,bold=True,bg=DARK,color=WHITE,size=12,align="center")
    ws1.row_dimensions[row].height = 26; row+=1

    ws1.merge_cells("A%d:H%d"%(row,row))
    c = ws1.cell(row=row,column=1,value=period)
    sty(c,bg=LBLUE,size=10,align="center")
    ws1.row_dimensions[row].height = 20; row+=1

    if account:
        ws1.merge_cells("A%d:H%d"%(row,row))
        c = ws1.cell(row=row,column=1,value=account)
        sty(c,size=9,align="center"); ws1.row_dimensions[row].height=18; row+=1

    if company and company not in ("nan",""):
        ws1.merge_cells("A%d:H%d"%(row,row))
        c = ws1.cell(row=row,column=1,value=company)
        sty(c,bold=True,bg=YELLOW,size=10,align="center"); ws1.row_dimensions[row].height=20; row+=1

    row+=1
    for i,(lbl,val,col) in enumerate([
        ("Boshlanish qoldighi",bal_start,DARK),
        ("Tugash qoldighi",bal_end,RED if bal_end<bal_start else GREEN)
    ]):
        c1=get_column_letter(1+i*4); c2=get_column_letter(3+i*4)
        ws1.merge_cells("%s%d:%s%d"%(c1,row,c2,row))
        c=ws1.cell(row=row,column=1+i*4,value=lbl)
        sty(c,bold=True,bg=col,color=WHITE,align="center"); ws1.row_dimensions[row].height=20
        ws1.merge_cells("%s%d:%s%d"%(c1,row+1,c2,row+1))
        c2c=ws1.cell(row=row+1,column=1+i*4,value=val)
        sty(c2c,bold=True,bg=LBLUE,align="center",size=13); c2c.number_format="#,##0.00"
        ws1.row_dimensions[row+1].height=24
    row+=3

    for ci,h in enumerate(["Sana","Vaqt","Hujjat","Kontragent","Chiqim","Kirim","Izoh"],1):
        c=ws1.cell(row=row,column=ci,value=h)
        sty(c,bold=True,bg=DARK,color=WHITE,align="center")
    ws1.row_dimensions[row].height=20; row+=1

    for tx in txs:
        bg="FFE8E8" if tx["debit"] else "E8FFE8"
        for ci,val in enumerate([tx["date"],tx["time"],tx["doc"],tx["counterparty"],tx["debit"],tx["credit"],tx["desc"]],1):
            c=ws1.cell(row=row,column=ci,value=val)
            sty(c,bg=bg,size=9)
            if ci==5 and val: sty(c,bold=True,bg=bg,color=RED,align="right",size=10); c.number_format="#,##0.00"
            if ci==6 and val: sty(c,bold=True,bg=bg,color=GREEN,align="right",size=10); c.number_format="#,##0.00"
        ws1.row_dimensions[row].height=16; row+=1

    row+=1
    for lbl,val,col,ci in [("Jami chiqim:",total_debit,RED,5),("Jami kirim:",total_credit,GREEN,6)]:
        ws1.merge_cells("A%d:D%d"%(row,row))
        c=ws1.cell(row=row,column=1,value=lbl); sty(c,bold=True,align="right",size=11)
        c2=ws1.cell(row=row,column=ci,value=val); sty(c2,bold=True,color=col,align="right",size=12)
        c2.number_format="#,##0.00"; ws1.row_dimensions[row].height=22; row+=1

    for i,w in enumerate([12,7,13,28,16,16,40],1):
        ws1.column_dimensions[get_column_letter(i)].width=w

    ws2 = wb.create_sheet("Firmalar")
    debit_by,credit_by = build_summary(txs)
    all_names = set(list(debit_by.keys())+list(credit_by.keys()))
    row2=1
    ws2.merge_cells("A%d:E%d"%(row2,row2))
    c=ws2.cell(row=row2,column=1,value="FIRMALAR BO'YICHA JAMLAMA")
    sty(c,bold=True,bg=DARK,color=WHITE,size=13,align="center"); ws2.row_dimensions[row2].height=28; row2+=1
    ws2.merge_cells("A%d:E%d"%(row2,row2))
    c=ws2.cell(row=row2,column=1,value=period)
    sty(c,bg=LBLUE,size=10,align="center"); ws2.row_dimensions[row2].height=20; row2+=2
    for ci,h in enumerate(["Firma nomi","Chiqim soni","Jami chiqim","Kirim soni","Jami kirim"],1):
        c=ws2.cell(row=row2,column=ci,value=h)
        sty(c,bold=True,bg=DARK,color=WHITE,align="center")
    ws2.row_dimensions[row2].height=22; row2+=1
    for name in sorted(all_names):
        d=debit_by.get(name,{"sum":0,"count":0}); cr=credit_by.get(name,{"sum":0,"count":0})
        bg="FFF0F0" if d["sum"]>cr["sum"] else "F0FFF0" if cr["sum"]>d["sum"] else "FFFFFF"
        for ci,val in enumerate([name,d["count"] or "",d["sum"] or "",cr["count"] or "",cr["sum"] or ""],1):
            c=ws2.cell(row=row2,column=ci,value=val); sty(c,bg=bg,size=10)
            if ci==3 and val: sty(c,bold=True,bg=bg,color=RED,align="right",size=11); c.number_format="#,##0.00"
            if ci==5 and val: sty(c,bold=True,bg=bg,color=GREEN,align="right",size=11); c.number_format="#,##0.00"
        ws2.row_dimensions[row2].height=20; row2+=1
    for col,w in zip(["A","B","C","D","E"],[35,13,18,13,18]):
        ws2.column_dimensions[col].width=w
    wb.save(output_path)

def build_message(info, txs, total_debit, total_credit):
    debit_by,credit_by = build_summary(txs)
    all_names = set(list(debit_by.keys())+list(credit_by.keys()))
    msg = []
    msg.append("*"+info.get("bank","")[:60]+"*")
    if info.get("period"): msg.append(info["period"][:80])
    if info.get("company","") not in ("","nan"): msg.append("Korxona: *"+info["company"][:50]+"*")
    msg.append("")
    msg.append("Boshlanish: `%s` sum" % "{:,.0f}".format(info.get("bal_start",0)))
    msg.append("Tugash: `%s` sum" % "{:,.0f}".format(info.get("bal_end",0)))
    msg.append("")
    msg.append("Jami chiqim: *%s sum*" % "{:,.0f}".format(total_debit))
    msg.append("Jami kirim: *%s sum*" % "{:,.0f}".format(total_credit))
    msg.append("")
    msg.append("*FIRMALAR (%d ta):*" % len(all_names))
    for name in sorted(all_names):
        d=debit_by.get(name,{"sum":0,"count":0}); cr=credit_by.get(name,{"sum":0,"count":0})
        msg.append("")
        msg.append("*"+name+"*")
        if d["sum"]: msg.append("  Chiqim: `%s` sum (%d ta)" % ("{:,.0f}".format(d["sum"]),d["count"]))
        if cr["sum"]: msg.append("  Kirim: `%s` sum (%d ta)" % ("{:,.0f}".format(cr["sum"]),cr["count"]))
    return "\n".join(msg)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Salom!\n\n.xlsb, .xlsx yoki .xls bank kochirmasini yuboring.\nXabar va Excel hisobot qaytaraman."
    )

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc: return
    fname = doc.file_name or ""
    if not (fname.lower().endswith(".xlsb") or fname.lower().endswith(".xlsx") or fname.lower().endswith(".xls")):
        await update.message.reply_text("Iltimos, .xlsb, .xlsx yoki .xls fayl yuboring.")
        return
    await update.message.reply_text("Fayl qayta ishlanmoqda...")
    with tempfile.TemporaryDirectory() as tmpdir:
        fpath = os.path.join(tmpdir, fname)
        file = await doc.get_file()
        await file.download_to_drive(fpath)
        try:
            info, txs, total_debit, total_credit = parse_file(fpath, fname)
            msg = build_message(info, txs, total_debit, total_credit)
            if len(msg) > 4000: msg = msg[:4000] + "\n..."
            await update.message.reply_text(msg, parse_mode="Markdown")
            out_path = os.path.join(tmpdir, "hisobot.xlsx")
            create_excel(info, txs, total_debit, total_credit, out_path)
            with open(out_path,"rb") as f:
                await update.message.reply_document(
                    document=f, filename="bank_hisoboti.xlsx",
                    caption="Excel hisoboti tayyor! (%d ta tranzaksiya)" % len(txs)
                )
        except Exception as e:
            await update.message.reply_text("Xato: " + str(e))

async def handle_other(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Iltimos, .xlsb, .xlsx yoki .xls fayl yuboring.")

def main():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_other))
    print("Bot ishlamoqda...")
    app.run_polling()

if __name__ == "__main__":
    main()
