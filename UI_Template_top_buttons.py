from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QPushButton, QHBoxLayout


class main_top_buttons(QHBoxLayout):
    ontopbuttonChanged = pyqtSignal(object)

    def __init__(self, parent, total_programs):
        super(main_top_buttons, self).__init__(parent)

        # program buttons passed from main UI file
        self.total_programs = total_programs
        self.top_buttons_UI()

    def top_buttons_UI(self):

        # add buttons for each program at top of UI, total_programs is a global var
        for i in self.total_programs:
            btn = QPushButton(i)
            btn.adjustSize()
            btn.setCursor(QCursor(Qt.PointingHandCursor))
            btn.clicked.connect(self.switch_leftsidebuttons_signal)
            self.addWidget(btn)

        # align buttons to the top/left on UI
        self.setAlignment(Qt.AlignLeft | Qt.AlignTop)


    def switch_leftsidebuttons_signal(self):
        # pass button height to main UI for setting maximum height the UI element can be moved to by the user

        self.ontopbuttonChanged.emit(self.sender())


if __name__ == "__main__":
    pass
