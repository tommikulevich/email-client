import sys
import email
import smtplib
import imaplib
import time
from email.mime.text import MIMEText
from PySide6.QtGui import QIcon
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import *

SMTP_SERVER = 'smtp.poczta.onet.pl'
SMTP_PORT = 587
IMAP_SERVER = 'imap.poczta.onet.pl'
IMAP_PORT = 993


class EmailClient(QWidget):
    def __init__(self, username, password, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Email Client')
        self.setWindowIcon(QIcon(self.style().standardPixmap(QStyle.SP_DirHomeIcon)))

        self.username = username
        self.password = password
        self.actualMessagesNum = self.getInitMessageNum()

        self.tabs = QTabWidget()
        self.newMailTab = QWidget()
        self.sentTab = QWidget()
        self.inboxTab = QWidget()

        self.sendLayout = QVBoxLayout()
        self.sentLayout = QVBoxLayout()
        self.inboxLayout = QVBoxLayout()

        self.toLabel = QLabel()
        self.toLabel.setText(f'<b>To:</b>')
        self.toField = QLineEdit('tommikulevich@gmail.com')
        self.subjectLabel = QLabel()
        self.subjectLabel.setText(f'<b>Subject:</b>')
        self.subjectField = QLineEdit('Attempt to send (educational purposes)')
        self.messageLabel = QLabel()
        self.messageLabel.setText(f'<b>Message:</b>')
        self.messageField = QTextEdit('WNO classes at Gdańsk University of Technology')
        self.sendButton = QPushButton('Send')

        self.sentList = QListWidget()
        self.inboxList = QListWidget()

        self.sendLayout.addWidget(self.toLabel)
        self.sendLayout.addWidget(self.toField)
        self.sendLayout.addWidget(self.subjectLabel)
        self.sendLayout.addWidget(self.subjectField)
        self.sendLayout.addWidget(self.messageLabel)
        self.sendLayout.addWidget(self.messageField)
        self.sendLayout.addWidget(self.sendButton)

        self.newMailTab.setLayout(self.sendLayout)
        self.sentLayout.addWidget(self.sentList)
        self.sentTab.setLayout(self.sentLayout)
        self.inboxLayout.addWidget(self.inboxList)
        self.inboxTab.setLayout(self.inboxLayout)

        self.tabs.addTab(self.inboxTab, 'Inbox')
        self.tabs.addTab(self.sentTab, 'Sent')
        self.tabs.addTab(self.newMailTab, 'New email')

        self.keywordLabel = QLabel('Find:')
        self.keywordField = QLineEdit()
        self.keywordField.editingFinished.connect(self.refreshInboxList)
        self.inboxLayout.addWidget(self.keywordLabel)
        self.inboxLayout.addWidget(self.keywordField)

        mainLayout = QVBoxLayout()
        mainLayout.addWidget(self.tabs)
        self.setLayout(mainLayout)

        self.timer = QTimer()
        self.startTimer()

        self.sendButton.clicked.connect(self.sendEmail)
        self.sentList.itemDoubleClicked.connect(self.showEmail)
        self.inboxList.itemDoubleClicked.connect(self.showEmail)
        self.refreshSentList()
        self.refreshInboxList()

    # Count messages when user starts client
    def getInitMessageNum(self):
        with imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT) as serverIMAP:
            serverIMAP.login(self.username, self.password)
            serverIMAP.select('inbox')
            status, messages = serverIMAP.search(None, 'ALL')
            return len(messages[0].split())

    # Timer to refresh email lists
    def startTimer(self):
        self.timer.timeout.connect(self.refreshEmailLists)
        self.timer.start(30000)  # Update email lists every 30 sec

    # Refresh email lists
    def refreshEmailLists(self):
        self.refreshInboxList()
        self.refreshSentList()

    # Get sent emails using IMAP
    def refreshSentList(self):
        with imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT) as serverIMAP:
            self.sentList.clear()

            serverIMAP.login(self.username, self.password)
            serverIMAP.select('Sent')
            status, messages = serverIMAP.search(None, 'ALL')

            items = []
            for num in messages[0].split():
                status, message = serverIMAP.fetch(num, '(RFC822)')
                emailMessage = email.message_from_bytes(message[0][1])
                subject = emailMessage['Subject']
                subject = self.decodeSubject(subject)

                item = QListWidgetItem(f'{subject}')
                item.setIcon(self.style().standardIcon(QStyle.SP_ArrowForward))
                item.email = emailMessage
                item.from_ = emailMessage['From']
                item.to = emailMessage['To']

                items.append(item)

            for item in items[::-1]:
                self.sentList.addItem(item)

    # Get inbox emails using IMAP
    def refreshInboxList(self):
        with imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT) as serverIMAP:
            self.inboxList.clear()

            serverIMAP.login(self.username, self.password)
            serverIMAP.select('inbox')
            status, messages = serverIMAP.search(None, 'ALL')
            keyword = self.keywordField.text().lower()  # Get keyword from field

            # Check if new messages have arrived, then send autoresponse
            newMessagesNum = len(messages[0].split())
            if newMessagesNum > self.actualMessagesNum:
                self.sendAutoresponse(messages, serverIMAP)
                self.actualMessagesNum = newMessagesNum  # Update the number of actual messages

            items = []
            for num in messages[0].split():
                status, message = serverIMAP.fetch(num, '(RFC822)')
                emailMessage = email.message_from_bytes(message[0][1])
                subject = emailMessage['Subject']
                subject = self.decodeSubject(subject)

                if keyword in subject.lower():  # Search by keyword
                    item = QListWidgetItem(f'{subject}')
                    item.setIcon(self.style().standardIcon(QStyle.SP_ArrowForward))
                    item.email = emailMessage
                    item.from_ = emailMessage['From']
                    item.to = emailMessage['To']

                    items.append(item)

            for item in items[::-1]:
                self.inboxList.addItem(item)

    # Send autoresponse
    def sendAutoresponse(self, messages, serverIMAP):
        for num in messages[0].split()[self.actualMessagesNum:]:
            status, message = serverIMAP.fetch(num, '(RFC822)')
            emailMessage = email.message_from_bytes(message[0][1])
            sender = email.utils.parseaddr(emailMessage['From'])[1]

            message = MIMEText("For edu purposes", _charset='utf-8')
            message['Subject'] = 'Autoresponse: Thank you for your email!'
            message['From'] = self.username
            message['To'] = sender

            try:
                print("Connecting to server...")
                with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as serverSMPT:
                    serverSMPT.starttls()
                    print("Trying to login...")
                    serverSMPT.login(self.username, self.password)
                    print("Sending email...")
                    serverSMPT.send_message(message)
            except smtplib.SMTPAuthenticationError as e:
                print(f'SMTP Authentication Error: {e}')
            except smtplib.SMTPConnectError as e:
                print(f'SMTP Connection Error: {e}')
            except smtplib.SMTPException as e:
                print(f'SMTP Exception: {e}')
            else:
                print('Autoresponse sent successfully to:', sender)

    # Show email in new window
    def showEmail(self, item):
        emailMessage = item.email
        from_ = emailMessage['From']
        to = emailMessage['To']
        subject = emailMessage['Subject']
        subject = self.decodeSubject(subject)
        body = ''

        if emailMessage.is_multipart():
            for part in emailMessage.walk():
                contentType = part.get_content_type()
                contentDisp = str(part.get("Content-Disposition"))

                if "attachment" not in contentDisp:
                    if "text/plain" in contentType:
                        body = part.get_payload(decode=True).decode()
                    elif "text/html" in contentType:
                        body = part.get_payload(decode=True).decode()

        else:
            body = emailMessage.get_payload(decode=True).decode()

        window = QDialog()
        window.setWindowTitle(subject)
        window.setWindowIcon(QIcon(self.style().standardPixmap(QStyle.SP_FileDialogContentsView)))

        fromLabel = QLabel()
        fromLabel.setText(f'<b>From:</b> {from_}')
        toLabel = QLabel()
        toLabel.setText(f'<b>To:</b> {to}')
        subjectLabel = QLabel()
        subjectLabel.setText(f'<b>Subject:</b> {subject}')
        bodyLabel = QLabel(body)

        layout = QVBoxLayout()
        layout.addWidget(fromLabel)
        layout.addWidget(toLabel)
        layout.addWidget(subjectLabel)
        layout.addWidget(bodyLabel)
        window.setLayout(layout)

        window.exec()

    # Decode subject (UTF-8)
    @staticmethod
    def decodeSubject(subject):
        decoded = ''.join(text if isinstance(text, str) else text.decode(charset or 'utf-8')
                          for text, charset in email.header.decode_header(subject))
        return decoded

    # Send message using SMTP
    def sendEmail(self):
        # Email details from GUI fields
        message = MIMEText(self.messageField.toPlainText(), _charset='utf-8')
        message['Subject'] = self.subjectField.text()
        message['From'] = self.username
        message['To'] = self.toField.text()

        # Connect, login and send with SMTP
        try:
            print("Connecting to server...")
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as serverSMPT:
                serverSMPT.starttls()
                print("Trying to login...")
                serverSMPT.login(self.username, self.password)
                print("Sending email...")
                serverSMPT.send_message(message)
        except smtplib.SMTPAuthenticationError as e:
            print(f'SMTP Authentication Error: {e}')
        except smtplib.SMTPConnectError as e:
            print(f'SMTP Connection Error: {e}')
        except smtplib.SMTPException as e:
            print(f'SMTP Exception: {e}')
        else:
            print('Email sent successfully!')

            try:
                with imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT) as serverIMAP:
                    serverIMAP.login(self.username, self.password)
                    serverIMAP.append('Sent', None, imaplib.Time2Internaldate(time.time()),
                                      str(message).encode('utf-8'))
                    print("Added to Sent")
            except imaplib.IMAP4.error as e:
                print(f'IMAP Error: {e}')

        # Refresh sent list
        self.refreshSentList()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    client = EmailClient(username='tmq-contact@op.pl', password='SL4bQr.w$Rq!Pj@')
    client.show()
    sys.exit(app.exec())
