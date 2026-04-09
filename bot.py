from __future__ import annotations

import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Any

from telegram import ReplyKeyboardMarkup, ReplyKeyboardRemove, Update
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

from app.config import settings, validate_settings
from app.google_services import DriveRepository, SheetRepository, build_clients, staff_to_json
from app.template_renderer import create_placeholder_template, fill_template, convert_docx_to_pdf

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

(
    DOC_NO,
    DOC_DATE,
    RECIPIENT,
    SITE_CODE,
    WORK_DETAIL,
    WORK_DATE,
    WORK_TIME,
    STAFF_INPUT,
    CONFIRM,
) = range(9)


def _request_id() -> str:
    return datetime.now().strftime("REQ-%Y%m%d-%H%M%S")


def _user_label(update: Update) -> str:
    user = update.effective_user
    if not user:
        return "Unknown"
    name = " ".join(filter(None, [user.first_name, user.last_name])).strip()
    return name or user.username or str(user.id)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "พร้อมใช้งานครับ\n\nใช้ /newsite เพื่อสร้างหนังสือแจ้งเข้าไซต์\nใช้ /history เพื่อดูประวัติล่าสุด"
    )


async def newsite(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    context.user_data["request_id"] = _request_id()
    context.user_data["staff_items"] = []
    await update.message.reply_text("กรอกเลขที่หนังสือ เช่น North2026_NOR1-0117")
    return DOC_NO


async def collect_doc_no(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["document_no"] = update.message.text.strip()
    await update.message.reply_text("วันที่หนังสือ เช่น 01 เมษายน 2569")
    return DOC_DATE


async def collect_doc_date(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["doc_date_th"] = update.message.text.strip()
    await update.message.reply_text("เรียน / ผู้รับ เช่น เจ้าของอาคาร / เจ้าของพื้นที่")
    return RECIPIENT


async def collect_recipient(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["recipient"] = update.message.text.strip()
    await update.message.reply_text("รหัสไซต์ เช่น CMIA218")
    return SITE_CODE


async def collect_site_code(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["site_code"] = update.message.text.strip()
    await update.message.reply_text("รายละเอียดงาน เช่น ตรวจสอบแก้ไขระบบไฟฟ้า")
    return WORK_DETAIL


async def collect_work_detail(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["work_detail"] = update.message.text.strip()
    await update.message.reply_text("วันที่เข้าพื้นที่ เช่น 01 – 03 เมษายน 2569")
    return WORK_DATE


async def collect_work_date(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["work_date"] = update.message.text.strip()
    await update.message.reply_text("เวลาเข้าพื้นที่ เช่น 08:00 – 17:30 น.")
    return WORK_TIME


async def collect_work_time(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["work_time"] = update.message.text.strip()
    await update.message.reply_text(
        "กรอกรายชื่อเจ้าหน้าที่ทีละคนในรูปแบบ\nชื่อ|เบอร์โทร|บทบาท\n\nตัวอย่าง\nนายจตุพล จันทร์ทอง|081-595-2897|ผู้ดำเนินงาน BBTEC\n\nเมื่อครบแล้วพิมพ์ `done`",
        parse_mode="Markdown",
    )
    return STAFF_INPUT


async def collect_staff(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    if text.lower() == "done":
        if not context.user_data["staff_items"]:
            await update.message.reply_text("กรุณาใส่เจ้าหน้าที่อย่างน้อย 1 คน")
            return STAFF_INPUT
        summary = _build_summary(context.user_data)
        await update.message.reply_text(
            summary,
            reply_markup=ReplyKeyboardMarkup([["ยืนยัน", "ยกเลิก"]], resize_keyboard=True),
        )
        return CONFIRM

    parts = [p.strip() for p in text.split("|")]
    if len(parts) != 3:
        await update.message.reply_text("รูปแบบไม่ถูกต้อง กรุณาใช้ ชื่อ|เบอร์โทร|บทบาท")
        return STAFF_INPUT
    context.user_data["staff_items"].append({"name": parts[0], "phone": parts[1], "role": parts[2]})
    await update.message.reply_text(f"เพิ่มแล้ว {parts[0]}\nเพิ่มคนถัดไปได้เลย หรือพิมพ์ done")
    return STAFF_INPUT


async def confirm(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text.strip()
    if text == "ยกเลิก":
        await update.message.reply_text("ยกเลิกรายการแล้ว", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END
    if text != "ยืนยัน":
        await update.message.reply_text("กรุณาเลือก ยืนยัน หรือ ยกเลิก")
        return CONFIRM

    await update.message.reply_text("กำลังสร้างเอกสาร กรุณารอสักครู่...", reply_markup=ReplyKeyboardRemove())
    clients = context.application.bot_data["google_clients"]
    sheets = SheetRepository(clients.spreadsheet)
    drive = DriveRepository(clients.drive, settings.google_drive_output_folder_id)

    request_id = context.user_data["request_id"]
    user = update.effective_user
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    payload = {
        "request_id": request_id,
        "created_at": now_str,
        "updated_at": now_str,
        "status": "draft",
        "telegram_user_id": user.id,
        "telegram_name": _user_label(update),
        "telegram_username": user.username or "",
        "chat_id": update.effective_chat.id,
        "document_no": context.user_data["document_no"],
        "doc_date_th": context.user_data["doc_date_th"],
        "recipient": context.user_data["recipient"],
        "site_code": context.user_data["site_code"],
        "work_detail": context.user_data["work_detail"],
        "work_date": context.user_data["work_date"],
        "work_time": context.user_data["work_time"],
        "staff_json": staff_to_json(context.user_data["staff_items"]),
        "request_summary": f"{context.user_data['site_code']} {context.user_data['work_detail']}",
    }
    sheets.append_request(payload)

    workdir = Path("output") / request_id
    workdir.mkdir(parents=True, exist_ok=True)

    placeholder_template = workdir / "site_letter_template_placeholder.docx"
    create_placeholder_template(settings.template_docx_path, str(placeholder_template))

    output_docx = workdir / f"{context.user_data['document_no']} ({context.user_data['site_code']}).docx"
    fill_template(
        str(placeholder_template),
        str(output_docx),
        {
            "DOC_NO": context.user_data["document_no"],
            "DOC_DATE_THAI": context.user_data["doc_date_th"],
            "RECIPIENT": context.user_data["recipient"],
            "SITE_CODE": context.user_data["site_code"],
            "WORK_DETAIL": context.user_data["work_detail"],
            "WORK_DATE": context.user_data["work_date"],
            "WORK_TIME": context.user_data["work_time"],
        },
        context.user_data["staff_items"],
    )
    output_pdf = convert_docx_to_pdf(str(output_docx), str(workdir))

    docx_info = drive.upload_file(str(output_docx), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    pdf_info = drive.upload_file(str(output_pdf), "application/pdf")

    sheets.update_request_by_id(
        request_id,
        {
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "status": "completed",
            "docx_drive_file_id": docx_info["file_id"],
            "docx_drive_link": docx_info["web_link"],
            "pdf_drive_file_id": pdf_info["file_id"],
            "pdf_drive_link": pdf_info["web_link"],
        },
    )

    with open(output_pdf, "rb") as f:
        await update.message.reply_document(document=f, filename=os.path.basename(output_pdf))
    await update.message.reply_text(
        f"สร้างเอกสารสำเร็จ\nrequest_id: {request_id}\nPDF: {pdf_info['web_link']}\nDOCX: {docx_info['web_link']}"
    )
    return ConversationHandler.END


async def history(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    clients = context.application.bot_data["google_clients"]
    sheets = SheetRepository(clients.spreadsheet)
    rows = sheets.list_recent_by_user(update.effective_user.id, limit=5)
    if not rows:
        await update.message.reply_text("ยังไม่พบประวัติการสร้างเอกสาร")
        return
    lines = ["ประวัติล่าสุดของคุณ"]
    for row in rows:
        lines.append(
            f"- {row.get('created_at','')} | {row.get('document_no','-')} | {row.get('site_code','-')} | {row.get('status','-')}"
        )
    await update.message.reply_text("\n".join(lines))


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("ยกเลิกรายการแล้ว", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END


def _build_summary(data: dict[str, Any]) -> str:
    staff_lines = []
    for idx, item in enumerate(data["staff_items"], start=1):
        staff_lines.append(f"{idx}. {item['name']} | {item['phone']} | {item['role']}")
    return (
        "ตรวจสอบข้อมูลก่อนสร้างเอกสาร\n\n"
        f"เลขที่: {data['document_no']}\n"
        f"วันที่: {data['doc_date_th']}\n"
        f"เรียน: {data['recipient']}\n"
        f"ไซต์: {data['site_code']}\n"
        f"งาน: {data['work_detail']}\n"
        f"วันที่เข้าพื้นที่: {data['work_date']}\n"
        f"เวลา: {data['work_time']}\n\n"
        "เจ้าหน้าที่\n" + "\n".join(staff_lines)
    )


def main() -> None:
    validate_settings()
    clients = build_clients(settings.google_service_account_json, settings.google_sheet_id)

    application = Application.builder().token(settings.telegram_bot_token).build()
    application.bot_data["google_clients"] = clients

    conversation = ConversationHandler(
        entry_points=[CommandHandler("newsite", newsite)],
        states={
            DOC_NO: [MessageHandler(filters.TEXT & ~filters.COMMAND, collect_doc_no)],
            DOC_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, collect_doc_date)],
            RECIPIENT: [MessageHandler(filters.TEXT & ~filters.COMMAND, collect_recipient)],
            SITE_CODE: [MessageHandler(filters.TEXT & ~filters.COMMAND, collect_site_code)],
            WORK_DETAIL: [MessageHandler(filters.TEXT & ~filters.COMMAND, collect_work_detail)],
            WORK_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, collect_work_date)],
            WORK_TIME: [MessageHandler(filters.TEXT & ~filters.COMMAND, collect_work_time)],
            STAFF_INPUT: [MessageHandler(filters.TEXT & ~filters.COMMAND, collect_staff)],
            CONFIRM: [MessageHandler(filters.TEXT & ~filters.COMMAND, confirm)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("history", history))
    application.add_handler(CommandHandler("cancel", cancel))
    application.add_handler(conversation)

    logger.info("Bot started")
    application.run_polling()


if __name__ == "__main__":
    main()
