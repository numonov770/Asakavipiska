import os
import logging
import tempfile
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, CommandHandler, filters, ContextTypes
import subprocess
import pandas as pd

TOKEN = os.environ.get("BOT_TOKEN", "YOUR_TOKEN_HERE")

logging.basicConfig(level=logging.INFO)

def parse_xlsb(filepath):
    # Convert xlsb to xlsx using LibreOffice
    out_dir = tempfile.mkdtemp()
    result = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "xlsx", "--outdir", out_dir, filepath],
        capture_output=True, text=True, timeout=30
    )
    xlsx_files = [f for f in os.listdir(out_dir) if f.endswith(".xlsx")]
    if not xlsx_files:
        raise Exception("Konvertatsiya muvaffaqiyatsiz")
    xlsx_path = os.path.join(out_dir, xlsx_files[0])
    df = pd.read_excel(xlsx_path, sheet_name=0, header=None)
    return df

def format_report(df):
    lines = []
    try:
        bank = str(df.iloc[0, 0]).strip()
        period_line = str(df.iloc[1, 0]).strip()
        account_line = str(df.iloc[2, 0]).replace('\xa0', ' ').strip()
        balance_start = str(df.iloc[3, 0]).strip()
        balance_end = str(df.iloc[3, 1]).strip()
        owner = str(df.iloc[3, 7]).strip() if pd.notna(df.iloc[3, 7]) else ""

        lines.append("🏦 *BANK KO'CHIRMASI*")
        lines.append(f"📍 {bank}")
        lines.append(f"📅 {period_line}")
        if owner and owner != "nan":
            lines.append(f"👤 Egasi: *{owner}*")
        lines.append("")
        lines.append(f"💰 {balance_start}")
        lines.append(f"💰 {balance_end}")
        lines.append("")

        # Find header row
        header_row = None
        for i, row in df.iterrows():
            if str(row.iloc[0]).strip() == "Дата":
                header_row = i
                break

        if header_row is None:
            lines.append("❌ Tranzaksiyalar topilmadi")
            return "\n".join(lines)

        # Transactions
        total_debit = 0
        total_credit = 0
        tx_lines = []

        for i in range(header_row + 1, len(df)):
            row = df.iloc[i]
            date_val = row.iloc[0]
            if pd.isna(date_val):
                continue
            date_str = str(date_val)
            if "Итого" in date_str or "Итоговый" in date_str:
                # totals row
                try:
                    td = float(row.iloc[5]) if pd.notna(row.iloc[5]) else 0
                    tc = float(row.iloc[6]) if pd.notna(row.iloc[6]) else 0
                    total_debit = td
                    total_credit = tc
                except:
                    pass
                continue

            try:
                if hasattr(date_val, 'strftime'):
                    date_str = date_val.strftime("%d.%m.%Y %H:%M")
                debit = float(row.iloc[5]) if pd.notna(row.iloc[5]) else None
                credit = float(row.iloc[6]) if pd.notna(row.iloc[6]) else None
                desc = str(row.iloc[7]).strip() if pd.notna(row.iloc[7]) else ""
                # Clean description
                if desc.startswith("00634"):
                    desc = desc[5:].strip()

                if debit:
                    tx_lines.append(f"🔴 *{date_str}*\n   ─ {int(debit):,} сўм чиқим\n   📝 {desc[:80]}")
                elif credit:
                    tx_lines.append(f"🟢 *{date_str}*\n   + {int(credit):,} сўм кирим\n   📝 {desc[:80]}")
            except Exception:
                continue

        lines.append(f"📊 *Tranzaksiyalar: {len(tx_lines)} ta*")
        lines.append("")
        lines.extend(tx_lines)
        lines.append("")
        lines.append(f"📤 Jami chiqim: *{int(total_debit):,} сўм*")
        lines.append(f"📥 Jami kirim: *{int(total_credit):,} сўм*")
    except Exception as e:
        lines.append(f"❌ Xato: {e}")
    return "\n".join(lines)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Salom!\n\n"
        "📁 Menga *.xlsb* yoki *.xlsx* bank ko'chirmasi faylini yuboring — "
        "uni o'qib, chiroyli hisobot qilib beraman! 🏦"
    )

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc:
        return

    fname = doc.file_name or ""
    if not (fname.endswith(".xlsb") or fname.endswith(".xlsx")):
        await update.message.reply_text("❗ Iltimos, .xlsb yoki .xlsx fayl yuboring.")
        return

    await update.message.reply_text("⏳ Fayl o'qilmoqda...")

    with tempfile.TemporaryDirectory() as tmpdir:
        fpath = os.path.join(tmpdir, fname)
        file = await doc.get_file()
        await file.download_to_drive(fpath)

        try:
            df = parse_xlsb(fpath)
            report = format_report(df)
            # Telegram 4096 char limit
            if len(report) > 4000:
                report = report[:4000] + "\n...(qisqartirildi)"
            await update.message.reply_text(report, parse_mode="Markdown")
        except Exception as e:
            await update.message.reply_text(f"❌ Xato yuz berdi: {e}")

async def handle_other(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📁 Iltimos, .xlsb yoki .xlsx fayl yuboring.")

def main():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_other))
    print("Bot ishlamoqda...")
    app.run_polling()

if __name__ == "__main__":
    main()
