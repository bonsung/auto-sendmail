"""
메일 자동발송 스크립트
- 파일별 수신자에게 엑셀 첨부 자동 발송
- 알리바바 기업메일 (SMTP SSL 465)
"""

import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime, date

# =============================================
# ⚙️  SMTP 설정
# =============================================
SMTP_SERVER = "smtp.qiye.aliyun.com"
SMTP_PORT   = 465
EMAIL_FROM  = "jp@easytech-qd.cn"
EMAIL_PASS  = "0EciggzKgxqOZXR6"
# =============================================

# =============================================
# ⚙️  발송 목록 (파일별 수신자 설정)
# =============================================
MAIL_LIST = [
    {
        "file"   : r"C:\Users\Samsung\Desktop\0업무\0 C3 2EH NB Price\1 알콜 중국 내수가격_HN.xlsx",
        "to"     : "yjwon27@hannong.co.kr",
        "to_name": "윤주원 팀장",
        "cc"     : ["qxy@easytech-qd.cn", "jppark0307@daum.net"],
        "bcc"    : ["mm-jordan@kyungsan08.com"],
    },
    {
        "file"   : r"C:\Users\Samsung\Desktop\0업무\0 C3 2EH NB Price\2 L_중국 NBA 수출가격.xlsx",
        "to"     : "sh-kang@lotte.net",
        "to_name": "강상혁 수석",
        "cc"     : ["qxy@easytech-qd.cn", "jppark0307@daum.net"],
        "bcc"    : ["yeon422@naver.com"],
    },
]
# =============================================


def send_one_mail(mail_info: dict, today: date):
    """한 건 발송"""
    today_str    = today.strftime("%Y년 %m월 %d일")
    today_suffix = today.strftime("%Y%m%d")
    filename     = os.path.basename(mail_info["file"])

    subject = f"Daily price_{today_suffix}"
    body    = f"""안녕하세요.

{today_str} 기준 중국 내수가격 자료를 첨부드립니다.

감사합니다."""

    msg = MIMEMultipart()
    msg["From"]    = EMAIL_FROM
    msg["To"]      = mail_info["to"]
    msg["CC"]      = ", ".join(mail_info["cc"])
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain", "utf-8"))

    # 첨부파일
    if not os.path.exists(mail_info["file"]):
        raise FileNotFoundError(f"첨부 파일 없음: {mail_info['file']}")

    with open(mail_info["file"], "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{filename}"'
        )
        msg.attach(part)

    # 수신자 전체 (TO + CC + BCC)
    recipients = [mail_info["to"]] + mail_info["cc"] + mail_info["bcc"]

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
        server.login(EMAIL_FROM, EMAIL_PASS)
        server.sendmail(EMAIL_FROM, recipients, msg.as_bytes())


def main():
    today = date.today()
    print("=" * 45)
    print("  중국 내수가격 메일 자동발송")
    print(f"  실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 45)

    for i, mail_info in enumerate(MAIL_LIST, 1):
        filename = os.path.basename(mail_info["file"])
        print(f"\n[{i}/{len(MAIL_LIST)}] 발송 중...")
        print(f"   수신: {mail_info['to_name']} ({mail_info['to']})")
        print(f"   참조: {', '.join(mail_info['cc'])}")
        print(f"   숨참: {', '.join(mail_info['bcc'])}")
        print(f"   첨부: {filename}")
        try:
            send_one_mail(mail_info, today)
            print(f"   ✅ 발송 완료!")
        except Exception as e:
            print(f"   ❌ 오류: {e}")

    print(f"\n{'=' * 45}")
    print(f"  전체 완료: {datetime.now().strftime('%H:%M:%S')}")
    print(f"{'=' * 45}")
    input("\n종료하려면 Enter를 누르세요.")


if __name__ == "__main__":
    main()
