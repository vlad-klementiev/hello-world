from fuzzywuzzy import fuzz
from TypingTest_Interface import *
from PyQt5 import QtWidgets
import sys
import threading
import time


class Functions:

    def read_excel_cells(self, fname, n, nsheet):
        import xlrd
        workbook = xlrd.open_workbook(fname)
        sheet = workbook.sheet_by_index(nsheet)
        cells = []

        for c in range(sheet.nrows):
            cell = str(sheet.cell_value(c, n))
            cells.append(cell)

        return cells

    def read_excel_cells_couple(self, fname, n1, n2):
        import xlrd
        workbook = xlrd.open_workbook(fname)
        sheet = workbook.sheet_by_index(0)
        cells = []
        for c in range(sheet.nrows):
            cell = []
            cell1 = str(sheet.cell_value(c, n1))
            cell2 = str(sheet.cell_value(c, n2))
            cell.append(cell1)
            cell.append(cell2)
            cells.append(cell)

        return cells


class MainWindow(QtWidgets.QMainWindow):

    # Variables section
    text_filename = r'Library.xlsx'
    headers = Functions().read_excel_cells(text_filename, 0, 0)
    text_collection = Functions().read_excel_cells_couple(text_filename, 0, 1)
    selected_header = '(Choose a text to train with)'
    session_duration = 60  # value by default
    target_text = ''
    stopSession = False
    # \end\ Variables section

    # Initialization function
    def __init__(self, parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Filling drop-down lists with content
        self.ui.comboBox.addItems(self.headers)
        self.ui.comboBox.activated[str].connect(self.onHeaderActivated)
        time_intervals = ['Easy (60s)', 'Medium (45s)', 'Hard (30s)']
        self.ui.comboBox_2.addItems(time_intervals)
        self.ui.comboBox_2.activated[str].connect(self.onTimeActivated)

        # Attaching functions to push-buttons
        self.ui.pushButton.clicked.connect(self.preStart)
        self.ui.pushButton_2.clicked.connect(self.stopSessionF)
        self.ui.label_2.clear()
        self.ui.label_7.clear()

    # Other functions

    # Defining what's happening when a user
    # chooses an option in the first drop-down list
    def onHeaderActivated(self, text):
        self.selected_header = text
        index = self.headers.index(self.selected_header)
        self.target_text = self.text_collection[index][1]
        self.ui.label_2.setText(self.target_text)
        self.ui.label_2.setWordWrap(True)

    # Defining what's happening when a user
    # chooses an option in the second drop-down list
    def onTimeActivated(self, text):
        if text == 'Easy (60s)':
            self.session_duration = 60
        elif text == 'Medium (45s)':
            self.session_duration = 45
        elif text == 'Hard (30s)':
            self.session_duration = 30

    def Session(self):
        # Launch a session only if a text is chosen
        if self.selected_header != '(Choose a text to train with)':

            # Initialization
            self.stopSession = False
            stop_time = time.time() + self.session_duration
            entered_text = ''

            # TIMER
            while True:
                # calculating remained time
                remained_time = round(stop_time - time.time())

                # displays actual time on the screen
                self.ui.label_5.setText(f'Timer: {remained_time}')
                self.ui.label_5.setWordWrap(True)

                time.sleep(1)  # retards loop's execution

                # The timer stops (and the session is terminated)
                # if one of three conditions is true
                if time.time() > stop_time:  # 1st
                    remained_time = 0
                    self.ui.label_5.setText(f'Timer: {remained_time}')
                    break
                elif entered_text == self.target_text or \
                        self.stopSession:  # 2nd & 3rd
                    break
                # Updating text which a user has entered to monitor the progress
                entered_text = self.ui.textEdit.toPlainText()
            # \end\ TIMER

            # Final reading of entered text from the textEdit
            entered_text = self.ui.textEdit.toPlainText()
            # Making the textEdit unavailable
            self.ui.textEdit.setEnabled(False)

            # CALCULATING SESSION'S RESULT
            if not self.stopSession:  # If a session is not interrupted

                # Calculating numeric results
                spent_time = self.session_duration - remained_time  # time that is spent on typing
                spent_time = spent_time / 60  # converting seconds into minutes
                speed = round(len(entered_text) / spent_time / 5)

                # Comparing entered and target texts
                text_similarity = fuzz.ratio(entered_text, self.target_text)

                # loading image
                if speed > 45:
                    if text_similarity > 70:
                        pixmap = QtGui.QPixmap(r"pics\Congrats.png").scaled(70, 70)
                        judgement = 'Great job, congrats! Your typing speed ' \
                                    'is above average, which is 30-45 WPM (words per min).'
                    else:
                        pixmap = QtGui.QPixmap(r"pics\Needs_improving.png").scaled(100, 100)
                        judgement = "Your typing speed is high enough, " \
                                    "but there are too many misprints :("

                elif speed > 30:
                    pixmap = QtGui.QPixmap(r"pics\Not_Bad.png").scaled(100, 100)
                    judgement = "Not bad! Your typing speed belongs to an " \
                                "average values' range, which is 30-45 WPM (words per min). "
                else:
                    pixmap = QtGui.QPixmap(r"pics\Needs_improving.png").scaled(100, 100)
                    judgement = "Your typing speed isn't too high. " \
                                "Normal speed should be at least 30 WPM (words per min). "

                # Showing result to a user
                result_message = f'Your speed is {speed} WPM.\n' \
                                 f'The text similarity is {text_similarity}%.\n' \
                                 f'Remained time is {remained_time}s.'
                self.ui.label_3.setText('\n\n\n'.join([result_message, judgement]))
                self.ui.label_3.setWordWrap(True)
                self.ui.label_7.setPixmap(pixmap)

            else:  # If a session is (!) interrupted
                self.ui.label_5.setText(f'Timer: 0')

    def preStart(self):
        # Turning the input field available,
        # clearing it and some labels before each session's start
        self.ui.textEdit.setEnabled(True)
        self.ui.label_3.clear()
        self.ui.label_7.clear()
        self.ui.textEdit.setPlainText("")

        # Start a session
        TH = threading.Thread(target=self.Session)
        TH.start()

    def stopSessionF(self):  # The function interrupts a session execution
        self.stopSession = True


# Entry point of the program
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    myapp = MainWindow()
    myapp.show()

    sys._excepthook = sys.excepthook

    def exception_hook(exctype, value, traceback):
        print(exctype, value, traceback)
        sys._excepthook(exctype, value, traceback)
        sys.exit(1)

    sys.excepthook = exception_hook
    sys.exit(app.exec_())
