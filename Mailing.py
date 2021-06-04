import smtplib
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase

def mailing(bot):
    sendEmail = "이메일 주소"
    recvEmail = "이메일주소"
    password = "비밀번호"

    smtpName = "smtp.naver.com"  # smtp 서버 주소
    smtpPort = 587  # smtp 포트 번호

    text = "본문 내용"
    msg = MIMEMultipart()


    msg['Subject'] = "제목"
    msg['From'] = sendEmail
    msg['To'] = recvEmail

    fileName = '파일경로'
    attachment = open(fileName,'rb')
    part = MIMEBase('application','octat-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition',"attachment; filename= " + fileName)
    msg.attach(part)

    s = smtplib.SMTP(smtpName, smtpPort)  # 메일 서버 연결
    s.starttls()  # TLS 보안 처리
    s.login(sendEmail, password)  # 로그인
    s.sendmail(sendEmail, recvEmail, msg.as_string())  # 메일 전송, 문자열로 변환하여 보냅니다.
    s.close()  # smtp 서버 연결을 종료합니다.
