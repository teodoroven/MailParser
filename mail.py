from modules import *

class Letter():

    def __init__(self, id: bytes, subject: str, sender: str, date: datetime, body: str):
        self.id: bytes = id
        self.subject: str = str(subject)
        self.sender: str = str(sender)
        self.date: datetime = date
        self.body = body
        
    def get(self, mail_part: Literal["subject", "address", "body", "date"]) -> str:
        match mail_part:
            case "subject":
                return self.subject
            case "body":
                return str(self.body)
            case "address":
                return self.sender
            case "date":
                return self.date.strftime(DATE_FORMAT_BASED)
            case _:
                raise TypeError(f"invalid {mail_part=}")

class IMAP():

    def __init__(self, server: str, address: str, password: str):
        self.imap_server = str(server)
        self.address = str(address)
        self.password = str(password)
        
        self.mail = None
        
    def start(self):
        self.mail = imaplib.IMAP4_SSL(self.imap_server)
        self.mail.login(self.address, self.password)
        self.mail.select("inbox")
    
    def send(self, address: str, subject: str, body: str):
        """
        Отправляет электронное письмо.

        :param to: Адрес получателя.
        :param subject: Тема письма.
        :param body: Текст письма.
        """
        msg = MIMEMultipart()
        msg["From"] = self.address
        msg["To"] = address
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "plain"))
        server = smtplib.SMTP_SSL("smtp.mail.ru", 465)
        server.login(self.address, self.password)
        server.sendmail(self.address, address, msg.as_string())
        server.quit()

    def get_letters(self, quantity: int, date: datetime = None, progress: list = None) -> list[Letter]:
        """
        Возвращает несколько последних писем с почты.

        :param quantity: Количество последних писем.
        :param date: Будут возвращены только письма после указанной даты.
        """
        
        self.start()
        
        result, data = self.mail.search(None, "ALL")
        mail_ids = data[0].split()
        messages = []
        
        date = date.replace(tzinfo=None)
        
        def process_message(msg_data):
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
            message_date = msg["Date"]
            print(message_date)
            message_date = message_date.replace(" (PDT)","")
            
            try:
                d: datetime = strptime(message_date)
            except Exception as err:
                # FIXME: Сообщение без даты больше не во входящих
                print("MAIL:",type(err),err,message_date)
                d = datetime.now()
            
            # FIXME: Проблемы с часовыми поясами!!!
            if d <= date:
                print(message_date, d, date)
            
                if type(progress) is list:
                    progress.append(False)
                    
                return
                
            sender = msg["From"]
            subject = msg["Subject"]
            
            # FIXME: проверить все составляющие сообщения
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                        break
            else:
                body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")
            
            subject = decode_header(subject)[0][0].decode()
            
            # FIXME: test!!!!
            
            messages.append(Letter(mail_id, subject, sender, d, body))
            
            if type(progress) is list:
                progress.append(True)
        
        for mail_id in mail_ids[-quantity:]:
            # FIXME: try_connection, чтобы не вылетало целиком
            result, msg_data = self.mail.fetch(mail_id, "(RFC822)")
            msg = msg_data
            process_message(msg)
            
            # TODO: Добавить многопоточность
            # def process_msg():
            #     return process_message(msg)
            
            # Thread(target=process_msg).start()
        
        self.mail.logout()
        return messages

        