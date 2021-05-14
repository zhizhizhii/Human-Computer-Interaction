from PyQt5 import QtCore, QtGui, QtWidgets, Qt
from PyQt5.QtGui import QMovie, QIcon
import win32api
import speech_recognition as sr
import os

class Ui_MainWindow(object):

    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(314, 550)
        MainWindow.setStyleSheet("background-color: rgb(0, 0, 0);")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(60, 280, 201, 51))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        font.setPointSize(12)
        self.label_3.setFont(font)
        self.label_3.setStyleSheet("color: rgb(0, 117, 210);")
        self.label_3.setWordWrap(True)
        self.label_3.setObjectName("label_3")

        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(60, 260, 201, 31))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        font.setPointSize(12)
        self.label_2.setFont(font)
        self.label_2.setStyleSheet("color: rgb(0, 117, 210);")
        self.label_2.setWordWrap(True)
        self.label_2.setObjectName("label_2")

        self.voiceFig = QtWidgets.QLabel(self.centralwidget)
        self.voiceFig.setGeometry(QtCore.QRect(70, 50, 161, 121))
        self.voiceFig.setText("")
        self.gif = QMovie("icon/voice.gif")
        self.voiceFig.setMovie(self.gif)
        self.gif.start()
        self.voiceFig.setScaledContents(True)
        self.voiceFig.setObjectName("voiceFig")

        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(70, 160, 180, 21))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        font.setPointSize(16)
        font.setBold(True);
        font.setStyleStrategy(QtGui.QFont.PreferAntialias)
        self.label.setFont(font)
        self.label.setStyleSheet("color: #fbbd08;bolder")
        self.label.setTextFormat(QtCore.Qt.AutoText)
        self.label.setWordWrap(True)
        self.label.setObjectName("label")

        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(60, 325, 201, 51))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        font.setPointSize(12)
        self.label_4.setFont(font)
        self.label_4.setStyleSheet("color: rgb(0, 117, 210);")
        self.label_4.setWordWrap(True)
        self.label_4.setObjectName("label_4")

        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(60, 370, 201, 51))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        font.setPointSize(12)
        self.label_5.setFont(font)
        self.label_5.setStyleSheet("color: rgb(0, 117, 210);")
        self.label_5.setWordWrap(True)
        self.label_5.setObjectName("label_5")

        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(60, 480, 201, 51))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        font.setPointSize(12)
        self.label_6.setFont(font)
        self.label_6.setStyleSheet("color: #f37b1d;")
        self.label_6.setWordWrap(True)
        self.label_6.setObjectName("label_4")

        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.pushButton = QtWidgets.QPushButton('', self.centralwidget)
        self.pushButton.resize(80, 80)
        self.pushButton.move(112, 420)
        self.pushButton.setStyleSheet("QPushButton{border-image: url(icon/play.gif)};")
        self.pushButton.setToolTip('Click Me to say something')
        self.pushButton.clicked.connect(self.clickedBtn)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Automatic Speech Recognition"))
        MainWindow.setWindowIcon(QIcon('./icon/icon.png'))
        self.label_3.setText(_translate("MainWindow", "1. Enjoy music by saying \"Music\""))
        self.label_2.setText(_translate("MainWindow", "You can:"))
        self.label.setText(_translate("MainWindow", "Hi! How can I help?"))
        self.label_4.setText(_translate("MainWindow", "2. Take some notes by saying \"Notepad\""))
        self.label_5.setText(_translate("MainWindow", "3. Take Baidu by saying \"Search\""))
        self.label_6.setText(_translate("MainWindow", "Click Me to say something!"))

    def clickedBtn(self):
        _translate = QtCore.QCoreApplication.translate
        self.label.setText(_translate("MainWindow", "I am Listening!"))
        guess = self.recognize_speech_from_mic(self.recognizer, self.microphone)
        if not guess['success']:
            QtWidgets.QMessageBox.critical(self, "错误", "系统错误!")
        if guess["error"]:
            print("ERROR: {}".format(guess["error"]))
            self.label.setText(_translate("MainWindow", "Error!"))
            return
        self.label.setText(_translate("MainWindow", guess["transcription"]))
        if guess["transcription"].lower()== "Music".lower():
            win32api.ShellExecute(0, 'open', os.getcwd()+'/music/music.mp3', '','',1)
        elif guess["transcription"].lower()== "Notepad".lower():
            win32api.ShellExecute(0, 'open', 'notepad.exe', '', '', 1)
        elif guess["transcription"].lower() == "Search".lower():
            win32api.ShellExecute(0, 'open', 'https://www.baidu.com', '', '', 1)
        else:
            win32api.ShellExecute(0, 'open', os.getcwd()+'/music/music.mp3', '','',1)
        self.label.setText(_translate("MainWindow", "Any other help?"))

    def recognize_speech_from_mic(self, recognizer, microphone):
        """Transcribe speech from recorded from `microphone`.

        Returns a dictionary with three keys:
        "success": a boolean indicating whether or not the API request was
                   successful
        "error":   `None` if no error occured, otherwise a string containing
                   an error message if the API could not be reached or
                   speech was unrecognizable
        "transcription": `None` if speech could not be transcribed,
                   otherwise a string containing the transcribed text
        """
        # check that recognizer and microphone arguments are appropriate type
        if not isinstance(recognizer, sr.Recognizer):
            raise TypeError("`recognizer` must be `Recognizer` instance")

        if not isinstance(microphone, sr.Microphone):
            raise TypeError("`microphone` must be `Microphone` instance")

        # adjust the recognizer sensitivity to ambient noise and record audio
        # from the microphone
        with microphone as source:
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source)

        # set up the response object
        response = {
            "success": True,
            "error": None,
            "transcription": None
        }

        # try recognizing the speech in the recording
        # if a RequestError or UnknownValueError exception is caught,
        #     update the response object accordingly
        try:
            response["transcription"] = recognizer.recognize_sphinx(audio)
            # response["transcription"] = recognizer.recognize_google(audio,language='zh-CN')
        except sr.RequestError:
            # API was unreachable or unresponsive
            response["success"] = False
            response["error"] = "API unavailable"
        except sr.UnknownValueError:
            # speech was unintelligible
            response["error"] = "Unable to recognize speech"

        return response