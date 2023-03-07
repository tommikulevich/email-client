import sys
import email
import smtplib
import imaplib
import time
from email.mime.text import MIMEText
from PySide6.QtWidgets import *

SMTP_SERVER = 'smtp.poczta.onet.pl'
SMTP_PORT = 587
IMAP_SERVER = 'imap.poczta.onet.pl'
IMAP_PORT = 993


class EmailClient(QWidget):
    def __init__(self, username, password, parent=None):
        super().__init__(parent)

        self.username = username
        self.password = password

        self.tabs = QTabWidget()
        self.newMailTab = QWidget()
        self.sentTab = QWidget()
        self.inboxTab = QWidget()

        self.sendLayout = QVBoxLayout()
        self.sentLayout = QVBoxLayout()
        self.inboxLayout = QVBoxLayout()

        self.toLabel = QLabel('To:')
        self.toField = QLineEdit('tommikulevich@gmail.com')
        self.subjectLabel = QLabel('Subject:')
        self.subjectField = QLineEdit('Attempt to send (educational purposes)')
        self.messageLabel = QLabel('Message:')
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

        self.sendButton.clicked.connect(self.sendEmail)
        self.refreshSentList()
        self.refreshInboxList()

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
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                print("Trying to login...")
                server.login(self.username, self.password)
                print("Sending email...")
                server.send_message(message)
        except smtplib.SMTPAuthenticationError as e:
            print(f'SMTP Authentication Error: {e}')
        except smtplib.SMTPConnectError as e:
            print(f'SMTP Connection Error: {e}')
        except smtplib.SMTPException as e:
            print(f'SMTP Exception: {e}')
        else:
            print('Email sent successfully!')

            try:
                with imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT) as server:
                    server.login(self.username, self.password)
                    server.append('Sent', None, imaplib.Time2Internaldate(time.time()), str(message).encode('utf-8'))
                    print("Added to Sent")
            except imaplib.IMAP4.error as e:
                print(f'IMAP Error: {e}')

        # Refresh sent list
        self.refreshSentList()

    # Get sent emails using IMAP (subjects only)
    def refreshSentList(self):
        with imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT) as server:
            server.login(self.username, self.password)
            server.select('Sent')
            status, messages = server.search(None, 'ALL')
            self.sentList.clear()

            for num in messages[0].split():
                status, message = server.fetch(num, '(RFC822)')
                emailMessage = email.message_from_bytes(message[0][1])
                subject = emailMessage['Subject']
                self.sentList.addItem(subject)

    # Get inbox emails using IMAP (subjects only)
    def refreshInboxList(self):
        with imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT) as server:
            server.login(self.username, self.password)
            server.select('inbox')
            status, messages = server.search(None, 'ALL')
            self.inboxList.clear()

            keyword = self.keywordField.text().lower()  # Get keyword from field

            for num in messages[0].split():
                status, message = server.fetch(num, '(RFC822)')
                emailMessage = email.message_from_bytes(message[0][1])
                subject = emailMessage['Subject']

                if keyword in subject.lower():  # Search by keyword
                    self.inboxList.addItem(subject)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    client = EmailClient(username='tmq-contact@op.pl', password='SL4bQr.w$Rq!Pj@')
    client.show()
    sys.exit(app.exec())
