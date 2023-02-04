from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QApplication
from datetime import datetime
import locale
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import pandas as pd
import glob
from datetime import datetime
import logging
from win32api import GetKeyState
from win32con import VK_CAPITAL
from unicode_tr import unicode_tr
import webbrowser
from datetime import timedelta

df2 = pd.DataFrame({"namesurname":[],"mail":[],"city":[],"date":[],"hour":[],"location":[]})
df3 = pd.DataFrame({"Mail":[], "Name Surname ":[], "Mail Check":[], "Time":[]})
df4 = pd.DataFrame({"namesurname":[],"mail":[],"city":[],"date":[],"hour":[],"location":[]})
days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
ccList = {
    "1":"xxx@gmail.com"
}

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class QTextEditLogger(logging.Handler):
    def __init__(self, parent):
        super().__init__()
        self.widget = QtWidgets.QPlainTextEdit(parent)
        self.widget.setReadOnly(True)

    def emit(self, record):
        msg = self.format(record)
        self.widget.appendPlainText(msg)

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(475, 360)
        Form.setMaximumSize(475, 360)
        Form.setMinimumSize(475, 360)

        Form.setWindowIcon(QtGui.QIcon(resource_path("icon.ico")))
        self.firstPixLabel = 10
        self.secondPixLabel = 270
        self.thirdPixLabel = 110
        self.fourthPixLabel = 450
        self.recentMailsHidden = True
        self.isDarkModeOn = False
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.getCapsLockState)
        self.timer.start(10)
        self.mode="-empty-"
        self.filePath=""
        self.user=""
        self.password=""
        self.lenght=0
        self.cc=""
        self.morefilenumber=0

        self.subject = "Mode 1 - Mail Subject"
        self.body = "Hi, Mr/Mrs {namesurname},\n\nCongratulations\n\nWe invite you for a face-to-face interview.\n\n Your face-to-face interview will be held in {city} on {day} at {hour}! \n\nInterviews that we will meet face-to-face with our candidates will not be compensated, and the candidates must be at the center where the interview will be held at least 15 minutes before the interview days and hours. \n\n{location}\n\nSee you at the interview."

        self.subject2 = "Mode 2 - Mail Subject"
        self.body2 = "Hi, Mr/Mrs {namesurname},\n\nCongratulations\n\nWe invite you for a face-to-face interview.\n\n Your face-to-face interview will be held in {city} on {day} at {hour}! \n\nInterviews that we will meet face-to-face with our candidates will not be compensated, and the candidates must be at the center where the interview will be held at least 15 minutes before the interview days and hours. \n\n{location}\n\nSee you at the interview."

        self.body3 = "Hi, Mr/Mrs {namesurname},\n\nCongratulations\n\nWe invite you for a face-to-face interview.\n\n Your face-to-face interview will be held in {city} on {day} at {hour}! \n\nInterviews that we will meet face-to-face with our candidates will not be compensated, and the candidates must be at the center where the interview will be held at least 15 minutes before the interview days and hours. \n\n{location}\n\nSee you at the interview."

        self.sentOrErrorLabel = QtWidgets.QLabel("", Form)
        self.sentOrErrorLabel.setGeometry(QtCore.QRect(170, 330, 300, 20))  
        self.sentOrErrorLabel.setObjectName("sentOrErrorLabel")
        self.sentOrErrorLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.sentOrErrorLabel.setText("-")

        self.attachedOrErrorLabel = QtWidgets.QLabel("", Form)
        self.attachedOrErrorLabel.setGeometry(QtCore.QRect(self.firstPixLabel, 130, 455, 20))
        self.attachedOrErrorLabel.setObjectName("attachedOrErrorLabel")
        self.attachedOrErrorLabel.setText("ATTACHED FİLE: -")

        self.runtimeLabel = QtWidgets.QPlainTextEdit(Form)
        self.runtimeLabel.setGeometry(QtCore.QRect(self.firstPixLabel, 260, 380, 60))
        self.runtimeLabel.setReadOnly(True)

        self.fileProcessorPushButton = QtWidgets.QPushButton(Form)
        self.fileProcessorPushButton.setGeometry(QtCore.QRect(395, 260, 70, 60))
        self.fileProcessorPushButton.setObjectName("fileProcessorPushButton")

        self.sendMailPushButton = QtWidgets.QPushButton(Form)
        self.sendMailPushButton.setGeometry(QtCore.QRect(self.firstPixLabel, 205, 455, 50))
        self.sendMailPushButton.setObjectName("sendMailPushButton")

        self.txtPathPushButton = QtWidgets.QPushButton(Form)
        self.txtPathPushButton.setGeometry(QtCore.QRect(self.firstPixLabel, 150, 455, 50))
        self.txtPathPushButton.setObjectName("txtPathPushButton")

        self.senderLineEdit = QtWidgets.QLineEdit(Form)
        self.senderLineEdit.setGeometry(QtCore.QRect(self.firstPixLabel, 20, 455, 25))
        self.senderLineEdit.setObjectName("senderLineEdit")
        self.senderLineEdit.setPlaceholderText("Sender")

        self.passwordLineEdit = QtWidgets.QLineEdit(Form)
        self.passwordLineEdit.setGeometry(QtCore.QRect(self.firstPixLabel, 50, 455, 25))
        self.passwordLineEdit.setObjectName("passwordLineEdit")
        self.passwordLineEdit.setPlaceholderText("Password")
        

        self.dateAndTimeLabel = QtWidgets.QLabel(Form)
        self.dateAndTimeLabel.setGeometry(QtCore.QRect(self.firstPixLabel, 330, 200, 20))
        self.dateAndTimeLabel.setObjectName("dateAndTimeLabel")
        self.dateAndTimeLabel.setAlignment(QtCore.Qt.AlignCenter)

        self.themeCheckBox = QtWidgets.QCheckBox(Form)
        self.themeCheckBox.setGeometry(QtCore.QRect(40, 85, 120, 20))
        self.themeCheckBox.setObjectName("themeCheckBox")

        self.showPasswordBox = QtWidgets.QCheckBox(Form)
        self.showPasswordBox.setGeometry(QtCore.QRect(170, 85, 120, 20))
        self.showPasswordBox.setObjectName("showPasswordBox")

        self.comboxBox = QtWidgets.QComboBox(Form)
        self.comboxBox.setGeometry(QtCore.QRect(310, 82, 120, 25))
        self.comboxBox.setObjectName("comboxBox")
        self.comboxBox.addItem("-empty-")
        self.comboxBox.addItem("Mode 1")
        self.comboxBox.addItem("Mode 2")
        self.comboxBox.addItem("Mode 3")

        self.sendMailDialog = QtWidgets.QDialog(Form)
        self.sendMailDialog.resize(400, 800)
        self.sendMailDialog.setWindowTitle("Mail Send Operation")
        self.sendMailDialog.setWindowModality(QtCore.Qt.ApplicationModal)
        self.sendMailDialog.setWindowFlag(QtCore.Qt.WindowContextHelpButtonHint, False)
        self.sendMailDialog.setMinimumSize(200,800)

        self.logTextBox = QTextEditLogger(self.sendMailDialog)
        self.logTextBox.setFormatter(logging.Formatter('%(levelname)s - %(message)s - %(asctime)s '))
        logging.getLogger().addHandler(self.logTextBox)
        logging.getLogger().setLevel(logging.NOTSET)

        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.logTextBox.widget)
        self.sendMailDialog.setLayout(layout)
        self.sendMailDialog.move(300, 100)

        self.capsLockStatePic = QtWidgets.QLabel(Form)
        self.capsLockStatePic.setGeometry(QtCore.QRect(self.fourthPixLabel, 50, 24, 24))
        self.capsLockStatePic.setPixmap(QtGui.QPixmap(resource_path("capsonimg.png")))
        self.capsLockStatePic.hide()

        self.comboxBox.activated[str].connect(self.onChanged)
        self.fileProcessorPushButton.clicked.connect(self.fileProcess)
        self.sendMailPushButton.clicked.connect(self.mailingOperation)
        self.txtPathPushButton.clicked.connect(self.attachFile)
        self.showPasswordBox.stateChanged.connect(self.showPwStateChanged)
        self.themeCheckBox.stateChanged.connect(self.themeCheckBoxStateChanged)
        self.passwordLineEdit.setEchoMode(QtWidgets.QLineEdit.Password)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        curLocale = locale.getlocale()
        locale.setlocale(locale.LC_TIME, curLocale)
        dateAndTime = datetime.now()
        currentDate = datetime.strftime(dateAndTime, "%D %X")

        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "| PyMailSender"))
        self.sendMailPushButton.setText(_translate("Form", "Send"))
        self.fileProcessorPushButton.setText(_translate("Form", "Process"))
        self.showPasswordBox.setText(_translate("Form", "Show Password"))
        self.dateAndTimeLabel.setText(_translate("Form", "Date: " + currentDate))
        self.themeCheckBox.setText(_translate("Form", "Dark Mode"))
        self.txtPathPushButton.setText(_translate("Form", "Attach File"))

    def setContent(self, content, level):
        QApplication.processEvents()
        if level == "info":
            logging.info(content)
        if level == "error":
            logging.error(content)

    def onChanged(self, text):
        """"This function use for set up auto mail and password according to mode"""
        self.emit("Set to {}.".format(text))
        self.mode = text
        if self.mode == "Mode 1":
            self.senderLineEdit.setText("xxxxx@gmail.com")
            self.passwordLineEdit.setText("xxxxxxxxx")# Mail Application Password
        if self.mode == "Mode 2" or self.mode == "Mode 3":
            self.senderLineEdit.setText("xxxxx@gmail.com")
            self.passwordLineEdit.setText("xxxxxxxxx")# Mail Application Password
        print(text)

    def excel_creator(self,namesurname,mail,city,date,hour,location,filePath):
        global df2
        df2 = df2.append({"namesurname":namesurname, "mail": mail, "city":city, "date":date, "hour":hour, "location":location}, ignore_index=True)
        df2.to_excel(filePath, index=False)
        self.emit(filePath+" named file has been created")
        # self.connect_functions(filePath)

    def connect_functions(self,filepath):#For use connect file process button to attach button
        self.filePath = filepath
        self.attachedOrErrorLabel.setText("ATTACHED FİLE: {}".format(filepath))
        webbrowser.open(filepath)

    def notsended_excel_creator(self,namesurname,mail,city,date,hour,location,filepath):
        """This function creates excel to check mail."""
        global df4
        df4 = df4.append({"namesurname":namesurname, "mail": mail, "city":city, "date":date, "hour":hour, "location":location}, ignore_index=True)
        df4.to_excel(self.filePath, index=False)
        self.emit("Unsent mails are saved in the file named filename+1.")

    def format_hour(self, given_date, n=5):
        formatted_date = datetime.strptime(str(given_date), '%H:%M:%S')
        final_time = formatted_date + timedelta(minutes=n)
        return formatted_date.strftime('%H:%M') + " - " + final_time.strftime('%H:%M')

    def fileProcess(self):
        """This function is used to clear the data to be sent from Excel and to be mailed."""
        if self.mode != "-empty-" or self.mode != "Mode 3":
            if self.mode == "Mode 1":
                print("Mode 1")
                filePath , check = QFileDialog.getOpenFileName(None, 'Open file','C:\\',"Excel files (*.xlsx)")
                try:
                    if check:
                        df = pd.read_excel(filePath)
                        for id,namesurname,mail,tel,etel,tc,hakem,city,date,hour,location,x,y,z in zip(df[df.columns[0]], df[df.columns[1]], df[df.columns[2]], df[df.columns[3]], df[df.columns[4]], df[df.columns[5]], df[df.columns[6]], df[df.columns[7]], df[df.columns[8]], df[df.columns[9]], df[df.columns[10]], df[df.columns[11]], df[df.columns[12]], df[df.columns[13]]):
                            if pd.isnull(date) or pd.isnull(hour) or pd.isnull(mail):
                                pass
                            else:
                                if date.day==1 or date.day==2 or date.day==3 or date.day==4 or date.day==5 or date.day==6 or date.day==7 or date.day==8 or date.day==9:
                                    day = "0" + str(date.day)
                                else:
                                    day = str(date.day)

                                if date.month==1 or date.month==2 or date.month==3 or date.month==4 or date.month==5 or date.month==6 or date.month==7 or date.month==8 or date.month==9:
                                    month = "0" + str(date.month)
                                else:
                                    month = str(date.month)

                                date = day + "." + month + "." + str(date.year) + " " + str(days[date.dayofweek])
                                city = unicode_tr(city).title()
                                if not location.startswith("Interview"):
                                    location = "Adress of Interview:\n"+location

                                    splitted=str(hour).split("-")
                                    if len(splitted) == 1:
                                        hour=self.format_hour(hour)

                                self.excel_creator(namesurname,mail,city,date,hour,location,filePath)
                    elif len(filePath)<5:
                        self.sentOrErrorLabel.setText("File path is too short.")
                    else:
                        self.sentOrErrorLabel.setText("Only .xlsx files allowed.")
                except FileNotFoundError:
                    self.sentOrErrorLabel.setText("File not found.")    
                except PermissionError:
                    self.sentOrErrorLabel.setText("Permission denied.")
            
            if self.mode == "Mode 2":
                print("Mode 2")
                filePath , check = QFileDialog.getOpenFileName(None, 'Open file','C:\\Users\\mucahitbektas\\Desktop',"Excel files (*.xlsx)")
                try:
                    if check:
                            df = pd.read_excel(filePath)    
                            cols = [0,1,2,4,5,7,8,12,13,14,15,16,18,19,20,21,22,23,24,25]
                            df.drop(df.columns[cols],axis=1,inplace=True)
                            for namesurname,mail,date,hour,location,city  in zip(df[df.columns[0]], df[df.columns[1]], df[df.columns[2]], df[df.columns[3]], df[df.columns[4]], df[df.columns[5]]):
                                if pd.isnull(date) or pd.isnull(hour) or pd.isnull(mail):
                                    pass
                                else:
                                    if date.day==1 or date.day==2 or date.day==3 or date.day==4 or date.day==5 or date.day==6 or date.day==7 or date.day==8 or date.day==9:
                                        day = "0" + str(date.day)
                                    else:
                                        day = str(date.day)

                                    if date.month==1 or date.month==2 or date.month==3 or date.month==4 or date.month==5 or date.month==6 or date.month==7 or date.month==8 or date.month==9:
                                        month = "0" + str(date.month)
                                    else:
                                        month = str(date.month)

                                    date = day + "." + month + "." + str(date.year) + " " + str(days[date.dayofweek])
                                    city = unicode_tr(city).title()
                                    if not location.startswith("Interview"):
                                        location = "Adress of Interview:\n"+location

                                    splitted=str(hour).split("-")
                                    if len(splitted) == 1:
                                        hour=self.format_hour(hour)

                                    self.excel_creator(namesurname,mail,city,date,hour,location,filePath)
                    elif len(filePath)<5:
                        self.sentOrErrorLabel.setText("File path is too short.")
                    else:
                        self.sentOrErrorLabel.setText("Only .xlsx files allowed.")
                except FileNotFoundError:
                    self.sentOrErrorLabel.setText("File not found.")    
                except PermissionError:
                    self.sentOrErrorLabel.setText("Permission denied.")
        else:
            self.emit("Please select mode.")

    def emit(self, record):
        self.runtimeLabel.appendPlainText(record)

    def mailingOperation(self):
        if self.mode != "-empty-":
            self.sendMailDialog.show()
            self.ccFinder()
            server = self.loginMail()
            self.oneHandler(server)
            self.quit(server)
        else:
            self.emit("Please select mode.")

    def ccFinder(self):
        df = pd.read_excel(self.filePath)
        if df.columns[2] == "city":
            city = str(df.iloc[0,2])
            self.cc = [ccList[i] for i in ccList.keys() if i == city][0]
            self.emit("{} is added to CC.".format(self.cc))
            self.setContent("{} is added to CC.".format(self.cc), "info")

            if self.mode == "Mode 3":
                self.cc = ""
                self.emit("CC is empty.")
                self.setContent("CC is empty.", "info")

    def loginMail(self):
        try:
            self.user = self.senderLineEdit.text()
            self.password = self.passwordLineEdit.text()
            server = smtplib.SMTP(host='smtp.gmail.com', port=587)
            server.starttls()
            server.login(self.user, self.password)
            print("--- LOGGED IN! --- ", self.user)
            self.setContent("--- LOGGED IN! ---", "info")
            self.emit("--- LOGGED IN! ---")
            return server

        except Exception as e:
            self.sentOrErrorLabel.setText(e)
            self.setContent(e, "error")

    def quit(self,server):
        try:
            server.quit()
            print("--- LOGGED OUT! --- ", self.user)
            self.setContent("--- LOGGED OUT! ---", "info")
            self.emit("--- LOGGED OUT! ---")
            self.sentOrErrorLabel.setText("Email sent to {} person.".format(self.lenght))
            
        except Exception as e:
            self.sentOrErrorLabel.setText(str(e))
            self.setContent(e, "error")

    def oneHandler(self,server):
        try:
            path = self.filePath.split(os.sep)
            file = path[-1]
            mainfile = "\\".join(path[:-1])
            files = glob.glob(os.path.join(mainfile, file))
            df = pd.read_excel(files[0])
            for namesurname, mail, city, day, hour, location in zip(df[df.columns[0]], df[df.columns[1]], df[df.columns[2]], df[df.columns[3]], df[df.columns[4]], df[df.columns[5]]):
                self.sendMail(namesurname, mail, city, day, hour, location,server)

        except Exception as e:
            self.sentOrErrorLabel.setText(str(e))
            self.setContent(e, "error")

    def sendMail(self,one,two,three,four,five,six,server):
        durum = False
        firstLen = len(self.senderLineEdit.text())
        secondLen = len(self.passwordLineEdit.text())

        if firstLen and secondLen != 0:
            if "@" and ".com" in two:
                try:
                    if self.mode == "Mode 1":
                        body = self.body.format(namesurname=one,city=three,day=four,hour=five,location=six)
                    if self.mode == "Mode 2":
                        body = self.body2.format(namesurname=one,city=three,day=four,hour=five,location=six)
                    if self.mode == "Mode 3":
                        body = self.body3.format(namesurname=one,day=four,hour=five,link=six)
                    msg = MIMEMultipart()
                    msg['From']=self.senderLineEdit.text()
                    msg['To']=two
                    if self.mode == "Mode 1":#For use set of subject according to mode
                        msg['Subject']=self.subject
                    if self.mode == "Mode 2" or self.mode == "Mode 3":
                        msg['Subject']=self.subject2
                    if self.mode == "Mode 1":#For use set of bcc according to mode
                        msg["Bcc"] = 'xxxx@gmail.com'
                    if self.mode == "Mode 2" or self.mode == "Mode 3":
                        msg["Bcc"] = 'xxxx@gmail.com'
                    msg["Cc"] = self.cc
                    msg.attach(MIMEText(body, 'plain'))
                    server.send_message(msg)
                    print(two , "'e mail gönderildi")
                    self.setContent(two, "info")
                    self.emit(two)
                    self.lenght+=1
                    durum=True
                    del msg

                except Exception as e:
                    print(e)
                    self.sentOrErrorLabel.setText(str(e))
                    self.setContent(e, "error")
                    # self.notsended_excel_creator(one,two,three,four,five,six) # for create not_sended excel

                finally:
                    global df3
                    df3 = df3.append({"Mail":two, "Name Surname":one, "Mail Check": durum, "Time":datetime.now()}, ignore_index=True)
                    os.chdir('\\'.join(self.filePath.split('/')[:-1]))
                    filename = self.filePath.split('/')[-1]
                    df3.to_excel("sended_{}.xlsx".format(filename), index=False)
            else:
                self.sentOrErrorLabel.setText("Please enter correct addresses.")
        else:
            self.sentOrErrorLabel.setText("Please fill all fields.")

    def showPwStateChanged(self):
        if self.showPasswordBox.isChecked():
            self.passwordLineEdit.setEchoMode(QtWidgets.QLineEdit.Normal)
        else:
            self.passwordLineEdit.setEchoMode(QtWidgets.QLineEdit.Password)

    def themeCheckBoxStateChanged(self):
        lightTheme = (open("themes\\lightTheme.qss", "r").read())
        darkTheme = (open("themes\\darkTheme.qss", "r").read())
        if self.themeCheckBox.isChecked():
            app.setStyleSheet(darkTheme)
            self.isDarkModeOn = True
        else:
            app.setStyleSheet(lightTheme)
            self.isDarkModeOn = False

    def getCapsLockState(self):
        cks = GetKeyState(VK_CAPITAL)
        
        if self.isDarkModeOn == False:
            self.capsLockStatePic.setPixmap(QtGui.QPixmap(resource_path(resource_path("capsonimg.png"))))
        else:
            self.capsLockStatePic.setPixmap(QtGui.QPixmap(resource_path(resource_path("capsonimgdark.png"))))
        
        if cks == 0:
            self.capsLockStatePic.hide()
        else:
            self.capsLockStatePic.show()
        
    def attachFile(self):
        if self.mode != "-empty-":
            filePath , check = QFileDialog.getOpenFileName(None, 'Open file','C:\\',"Excel files (*.xlsx)")
            try:
                if check:
                    with open(f'{filePath}', encoding="utf8"):
                        self.filePath = filePath
                        self.attachedOrErrorLabel.setText("ATTACHED FİLE: {}".format(filePath))
                        webbrowser.open(filePath)
                elif len(filePath)<5:
                    self.sentOrErrorLabel.setText("File path is too short.")
                else:
                    self.sentOrErrorLabel.setText("Only .xlsx files allowed.")
            except FileNotFoundError:
                self.sentOrErrorLabel.setText("File not found.")
            except PermissionError:
                self.sentOrErrorLabel.setText("Permission denied.")
        else:
            self.emit("Please select mode.")

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    app.setStyleSheet(open("themes\\lightTheme.qss", "r").read())
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())