# 📧 Email Client

> ☣ **Warning:** This project was created during my studies for educational purposes only. It may contain non-optimal or outdated solutions.

### 📑 About
**Email client** with GUI has been implemented. After starting program, a login window appears in which user enters his login and password, enters the parameters of IMAP and SMPT servers, and can turn on _the autoresponder_ for the selected period. Program receives emails (refreshes every 30 seconds) and allows user to view and send them. Also there is _a filter_ of received messages by keywords using intelligent word analysis (model used: sentence-transformers/paraphrase-multilingual-mpnet-base-v2).
**Created using:** PyCharm 2022.2.4 (Professional Edition) | Python 3.10.9 (WinPython 3.10 Release 2022-04)

### 📸 Screenshots

**Login window**
|User login details|Setting up servers|Autoresponder option|
|:-:|:-:|:-:|
|<img src="/_readmeImg/login_1.png?raw=true 'Login'">|<img src="/_readmeImg/login_2.png?raw=true 'Config'">|<img src="/_readmeImg/login_3.png?raw=true 'Autoresponder'">|


**Main window**

|Inbox view|Displaying emails|Section to send an email|
|:-:|:-:|:-:|
|<img src="/_readmeImg/main_1.png?raw=true 'Inbox'">|<img src="/_readmeImg/main_2.png?raw=true 'Displaying'">|<img src="/_readmeImg/main_3.png?raw=true 'Sending'">|


**Example of using smart filter** (words are marked by me)

<img src="/_readmeImg/filter.png?raw=true 'Filter'" width="400">
