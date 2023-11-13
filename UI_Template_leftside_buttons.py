from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QCursor, QFont
from PyQt5.QtWidgets import QWidget, QScrollArea, QVBoxLayout, QPushButton, QStackedLayout, QSizePolicy, QGroupBox


class main_leftside_buttons(QScrollArea):
    # signal emit for changes pages depending on which program UI is on when a leftside button on UI is hit
    programpageChanged = pyqtSignal(int)

    def __init__(self, parent, program_buttons):
        super(main_leftside_buttons, self).__init__(parent)

        # program buttons passed from main UI file
        self.program_buttons = program_buttons
        self.leftside_buttons_UI()

    def leftside_buttons_UI(self):

        # set size policy so that the leftside buttons UI can be scrunched as small as possible
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)

        # add upper level widget that contains layout to be added to scroll area
        self.leftside_buttons_UI_upperwidget = QWidget()

        # set resizable (otherwise widgets takes up a small area instead of entire layout, set scrollarea widget)
        self.setWidgetResizable(True)
        self.setWidget(self.leftside_buttons_UI_upperwidget)

        # create stacked layout for left side buttons of UI.  When button at top of program is clicked
        # it changes the stacked layout/buttons on left side of UI
        self.leftside_buttons_layout = QStackedLayout()

        # get a base width for buttons to set as the maximum width
        base_width = 0

        # populate stacked layout with buttons for each program from the program_buttons list of lists
        for i in self.program_buttons:
            group_box = QGroupBox()
            group_box.setStyleSheet("QGroupBox {border: 2px solid; background-color: lightgrey};")
            program_layout = QVBoxLayout()
            program_layout.setAlignment(Qt.AlignLeft | Qt.AlignTop)

            for k in i:
                btn = QPushButton(k)
                btn.adjustSize()
                btn.clicked.connect(self.switch_mainwork_layout_UI)
                btn.setFont(QFont('Times New Roman', 15))
                btn.setFlat(True)
                btn.setStyleSheet("QPushButton {color: magenta; text-align:left} QPushButton:hover {border: 1px solid gray}")
                btn.setCursor(QCursor(Qt.PointingHandCursor))

                btn_size = btn.frameGeometry().width()
                if btn_size > base_width:
                    base_width = btn_size

                program_layout.addWidget(btn)

            group_box.setLayout(program_layout)
            self.leftside_buttons_layout.addWidget(group_box)

        # add layout to upper level widget
        self.leftside_buttons_UI_upperwidget.setLayout(self.leftside_buttons_layout)

        # set maximum size window can be opened up , needed because this goes into a Qsplitter, which would
        # allow it to be expanded to any size.  Made it based buttons size (this may need to be changed
        # to have it based on something else
     #   self.setMaximumWidth(self.program1_button1.frameGeometry().width()+30)

        self.setMaximumWidth(base_width*2)

    # emit index and program of the leftside button that was hit, so that main UI can set correct pages of UI
    def switch_mainwork_layout_UI(self):
        # get index of which stacked layout of leftside buttons of UI is currently at
    #    current_program_index = self.leftside_buttons_layout.currentIndex()

        # this gets the index of button within it's corresponding layout.  (parent() finds the Groupbox,
        # findchild finds the layout withing groupbox, then indexof the button within layout
        sub_index = self.sender().parent().findChild(QVBoxLayout).indexOf(self.sender())

        # set mainwork layout to correct index based on what leftside button was clicked on UI
        self.programpageChanged.emit(sub_index)


if __name__ == "__main__":
    pass

