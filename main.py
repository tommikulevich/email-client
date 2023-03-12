import sys
import time
import ssl
import email
from email.mime.text import MIMEText
import smtplib
import imaplib
from bs4 import BeautifulSoup
from PySide6.QtGui import QIcon, Qt
from PySide6.QtCore import QTimer, QDate, QDateTime, QLocale
from PySide6.QtWidgets import *
from PySide6.QtWebEngineWidgets import QWebEngineView
from sentence_transformers import SentenceTransformer, util


class LoginWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Welcome')
        self.setWindowIcon(QIcon(self.style().standardPixmap(QStyle.SP_VistaShield)))
        self.setWindowModality(Qt.ApplicationModal)
        self.tabWidget = QTabWidget()

        self.loginTab = QWidget()
        self.emailLabel = QLabel('Email:')
        self.emailField = QLineEdit('tmq-contact@op.pl')
        self.passwordLabel = QLabel('Password:')
        self.passwordField = QLineEdit('SL4bQr.w$Rq!Pj@')
        self.passwordField.setEchoMode(QLineEdit.Password)

        layout = QGridLayout()
        layout.addWidget(self.emailLabel, 0, 0)
        layout.addWidget(self.emailField, 0, 1)
        layout.addWidget(self.passwordLabel, 1, 0)
        layout.addWidget(self.passwordField, 1, 1)
        self.loginTab.setLayout(layout)
        self.tabWidget.addTab(self.loginTab, 'Login')

        self.serverTab = QWidget()
        self.smtpLabel = QLabel('SMTP server:')
        self.smtpField = QLineEdit('smtp.poczta.onet.pl')   # smtp.student.pg.edu.pl
        self.smtpPortField = QLineEdit('465')               # 587/465
        self.imapLabel = QLabel('IMAP server:')
        self.imapField = QLineEdit('imap.poczta.onet.pl')   # imap.student.pg.edu.pl
        self.imapPortField = QLineEdit('993')               # 993

        layout = QGridLayout()
        layout.addWidget(self.smtpLabel, 0, 0)
        layout.addWidget(self.smtpField, 0, 1)
        layout.addWidget(self.smtpPortField, 0, 2)
        layout.addWidget(self.imapLabel, 1, 0)
        layout.addWidget(self.imapField, 1, 1)
        layout.addWidget(self.imapPortField, 1, 2)
        self.serverTab.setLayout(layout)
        self.tabWidget.addTab(self.serverTab, 'SMTP/IMAP Config')

        self.autoresponderTab = QWidget()
        self.autoresponderCheckbox = QCheckBox('Activate autoresponder')
        self.autoresponderCheckbox.setChecked(True)
        self.startLabel = QLabel('Start date:')
        self.startField = QDateEdit()
        self.startField.setCalendarPopup(True)
        self.startField.setDate(QDate.currentDate())
        self.endLabel = QLabel('End date:')
        self.endField = QDateEdit()
        self.endField.setCalendarPopup(True)
        self.endField.setDate(QDate.currentDate().addDays(7))

        layout = QGridLayout()
        layout.addWidget(self.autoresponderCheckbox, 0, 0, 1, 2)
        layout.addWidget(self.startLabel, 1, 0)
        layout.addWidget(self.startField, 1, 1)
        layout.addWidget(self.endLabel, 2, 0)
        layout.addWidget(self.endField, 2, 1)
        self.autoresponderTab.setLayout(layout)
        self.tabWidget.addTab(self.autoresponderTab, 'Autoresponder')

        self.loginButton = QPushButton('Log in')
        layout = QVBoxLayout()
        layout.addWidget(self.tabWidget)
        layout.addWidget(self.loginButton, alignment=Qt.AlignRight)
        self.setLayout(layout)

        self.loginButton.clicked.connect(self.accept)

    def getEmail(self):
        return self.emailField.text()

    def getPassword(self):
        return self.passwordField.text()

    def getSMTPServerInfo(self):
        return self.smtpField.text(), int(self.smtpPortField.text())

    def getIMAPServerInfo(self):
        return self.imapField.text(), int(self.imapPortField.text())


class EmailClient(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Email Client')
        self.setWindowIcon(QIcon(self.style().standardPixmap(QStyle.SP_DirHomeIcon)))
        self.model = SentenceTransformer('sentence-transformers/paraphrase-multilingual-mpnet-base-v2')

        self.loginWindow = LoginWindow()
        if self.loginWindow.exec() == QDialog.Accepted:
            self.username = self.loginWindow.getEmail()
            self.password = self.loginWindow.getPassword()
            self.smtpServer = self.loginWindow.getSMTPServerInfo()
            self.imapServer = self.loginWindow.getIMAPServerInfo()
        else:
            sys.exit(0)

        self.serverSMTP = smtplib.SMTP_SSL(self.smtpServer[0], self.smtpServer[1])
        self.serverSMTP.login(self.username, self.password)

        self.serverIMAP = imaplib.IMAP4_SSL(self.imapServer[0], self.imapServer[1],
                                            ssl_context=ssl.create_default_context().set_ciphers('DEFAULT'))
        self.serverIMAP.login(self.username, self.password)
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
        self.subjectField = QLineEdit('{My Email Client} Hello world!')
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

        self.logoutTab = QWidget()
        self.logoutLayout = QVBoxLayout()
        self.logoutButton = QPushButton('Logout')
        self.logoutButton.clicked.connect(self.logOut)
        self.logoutLayout.addWidget(self.logoutButton)
        self.logoutTab.setLayout(self.logoutLayout)

        self.tabs.addTab(self.inboxTab, 'Inbox')
        self.tabs.addTab(self.sentTab, 'Sent')
        self.tabs.addTab(self.newMailTab, 'New email')
        self.tabs.addTab(self.logoutTab, 'Logout')

        self.keywordLabel = QLabel()
        self.keywordLabel.setText(f'Filter:  <i>(e.g. greetings, astronomy, update)</i>')
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

    # Logout
    def logOut(self):
        self.serverSMTP.quit()
        self.serverIMAP.logout()
        self.close()

    # Count messages when user starts client
    def getInitMessageNum(self):
        self.serverIMAP.select('inbox')
        status, messages = self.serverIMAP.search(None, 'ALL')

        return len(messages[0].split())

    # Timer to refresh email lists
    def startTimer(self):
        self.timer.timeout.connect(self.refreshEmailLists)
        self.timer.start(30000)  # Update email lists every 20 sec

    # Refresh email lists
    def refreshEmailLists(self):
        print(f"{QLocale(QLocale.Polish).toString(QDateTime.currentDateTime().time(), 'hh:mm:ss')} Refreshing...")
        self.refreshInboxList()
        self.refreshSentList()

    # Get sent emails using IMAP
    def refreshSentList(self):
        self.sentList.clear()

        self.serverIMAP.select('Sent')
        status, messages = self.serverIMAP.search(None, 'ALL')

        items = []
        for num in messages[0].split():
            status, message = self.serverIMAP.fetch(num, '(RFC822)')
            emailMessage = email.message_from_bytes(message[0][1])
            subject = emailMessage['Subject']
            subject = self.decodeUTF8(subject)

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
        self.inboxList.clear()

        self.serverIMAP.select('inbox')
        status, messages = self.serverIMAP.search(None, 'ALL')
        keyword = self.keywordField.text()  # Get keyword from field

        # [Autoresponder: On] Check if new messages have arrived, then send autoresponse
        if self.loginWindow.autoresponderCheckbox.isChecked():
            newMessagesNum = len(messages[0].split())

        if newMessagesNum > self.actualMessagesNum:
            self.sendAutoresponse(messages)
            self.actualMessagesNum = newMessagesNum  # Update the number of actual messages

        for num in messages[0].split()[::-1]:
            status, message = self.serverIMAP.fetch(num, '(RFC822)')
            emailMessage = email.message_from_bytes(message[0][1])
            subject = emailMessage['Subject']
            subject = self.decodeUTF8(subject)

            # Intelligent filtering by keyword
            if keyword:
                keywordSentence = self.model.encode(keyword, convert_to_tensor=True)
                subjectSentence = self.model.encode(subject, convert_to_tensor=True)
                scoreCosSim = util.pytorch_cos_sim(keywordSentence, subjectSentence)
                score = scoreCosSim.item()
            else:
                score = 1

            if score > 0.4:
                item = QListWidgetItem(f'{subject}')
                item.setIcon(self.style().standardIcon(QStyle.SP_ArrowForward))
                item.email = emailMessage
                item.from_ = emailMessage['From']
                item.to = emailMessage['To']

                self.inboxList.addItem(item)

    # Send autoresponse
    def sendAutoresponse(self, messages):
        today = QDate.currentDate().toString('dd-MM-yyyy')
        startDate = self.loginWindow.startField.date().toPython().strftime('%d-%m-%Y')
        endDate = self.loginWindow.endField.date().toPython().strftime('%d-%m-%Y')

        if startDate <= today <= endDate:
            for num in messages[0].split()[self.actualMessagesNum:]:
                status, message = self.serverIMAP.fetch(num, '(RFC822)')
                emailMessage = email.message_from_bytes(message[0][1])
                sender = email.utils.parseaddr(emailMessage['From'])[1]

                message = MIMEText("Thank you for your email!", _charset='utf-8')
                message['Subject'] = f'[Autoresponse] Vacation: from {startDate} to {endDate}'
                message['From'] = self.username
                message['To'] = sender
                message.add_header('Disposition-Notification-To', self.username)
                message.add_header('Message-ID', email.utils.make_msgid())

                try:
                    print("Sending email...")
                    self.serverSMTP.send_message(message)
                except smtplib.SMTPException as e:
                    print(f'SMTP Exception: {e}')
                else:
                    print('Autoresponse sent successfully to:', sender)
        else:
            print('Autoresponse not sent! Today is outside provided date range')

    # Send message using SMTP
    def sendEmail(self):
        # Email details from GUI fields
        message = MIMEText(self.messageField.toPlainText(), _charset='utf-8')
        message['Subject'] = self.subjectField.text()
        message['From'] = self.username
        message['To'] = self.toField.text()
        message.add_header('Disposition-Notification-To', self.username)
        message.add_header('Message-ID', email.utils.make_msgid())

        # Send with SMTP
        try:
            print("Sending email...")
            self.serverSMTP.send_message(message)
        except smtplib.SMTPException as e:
            print(f'SMTP Exception: {e}')
        else:
            print('Email sent successfully!')

            try:
                self.serverIMAP.append('Sent', None, imaplib.Time2Internaldate(time.time()),
                                       str(message).encode('utf-8'))
            except imaplib.IMAP4.error as e:
                print(f'IMAP Error: {e}')

        # Refresh sent list
        self.refreshSentList()

    # Show email in new window
    def showEmail(self, item):
        emailMessage = item.email
        to = self.decodeUTF8(emailMessage['To'])
        from_ = self.decodeUTF8(email.utils.parseaddr(emailMessage['From'])[1])
        subject = self.decodeUTF8(emailMessage['Subject'])
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

                        try:
                            body = BeautifulSoup(body, 'html.parser').prettify()
                        except:
                            body = body.replace('\n', '<br>')
        else:
            body = emailMessage.get_payload(decode=True).decode()

            try:
                body = BeautifulSoup(body, 'html.parser').prettify()
            except:
                body = body.replace('\n', '<br>')

        window = QDialog()
        window.setWindowTitle(subject)
        window.setWindowIcon(QIcon(self.style().standardPixmap(QStyle.SP_FileDialogContentsView)))

        fromLabel = QLabel()
        fromLabel.setText(f'<b>From:</b> {from_}')
        toLabel = QLabel()
        toLabel.setText(f'<b>To:</b> {to}')
        subjectLabel = QLabel()
        subjectLabel.setText(f'<b>Subject:</b> {subject}')
        bodyLabel = QWebEngineView()
        bodyLabel.setHtml(body)

        layout = QVBoxLayout()
        layout.addWidget(fromLabel)
        layout.addWidget(toLabel)
        layout.addWidget(subjectLabel)
        layout.addWidget(bodyLabel)
        window.setLayout(layout)

        window.exec()

    # Decode subject/from/to (UTF-8)
    @staticmethod
    def decodeUTF8(subject):
        decoded = ''.join(text if isinstance(text, str) else text.decode(charset or 'utf-8')
                          for text, charset in email.header.decode_header(subject))
        return decoded


if __name__ == '__main__':
    app = QApplication(sys.argv)
    client = EmailClient()
    client.show()
    sys.exit(app.exec())
