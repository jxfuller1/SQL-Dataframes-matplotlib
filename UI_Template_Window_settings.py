from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QWidget, QScrollArea, QVBoxLayout, QPushButton, QGraphicsOpacityEffect, QHBoxLayout, QLabel, \
    QLineEdit, QCheckBox
import getpass
import os

# get computer user name, this is used for peoples custom setting initialization files
user = getpass.getuser()

# user main settings file name, note the "../" at the beginning must be there
# to work for relative path for relative location of the exe/python exe
user_settings = "\\main_settings_" + user + ".txt"

# get current working directory (needed if using on different computers
my_path = os.path.abspath(os.path.dirname(__file__))

# this would be the user settings path relative to exe location
user_path_settings = my_path + user_settings

# directory for default settings using relative path settings
default_path_settings = my_path + "\\main_settings.txt"


class window_settings(QScrollArea):
    # signal emit for changes pages depending on which program UI is on when a leftside button on UI is hit
    onsettingsChanged = pyqtSignal(int, int, bool, bool, bool, bool)

    def __init__(self, parent, windowsize_expansion_width, windowsize_expansion_height,
                 always_on_top, always_on_top_forever, left_window_collapse, top_window_collapse):
        super(window_settings, self).__init__(parent)

        # set variables passed down from main UI
        self.windowsize_expansion_width = windowsize_expansion_width
        self.windowsize_expansion_height = windowsize_expansion_height
        self.always_on_top = always_on_top
        self.always_on_top_forever = always_on_top_forever
        self.left_window_collapse = left_window_collapse
        self.top_window_collapse = top_window_collapse

        self.window_options()

    def window_options(self):

        # set so widgets take up entire frame geometry and disable horizontal scroll
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        # widget for slide out options window
        self.window_options_top_widget = QWidget()

        # make the slide out menu a little transparent by setting opacity, NOTE: not entirely
        # sure how to make this opacity not appear on child widget... so im setting the color of
        # the child widgets manually
        self.opacity = QGraphicsOpacityEffect()
        self.opacity.setOpacity(.95)
        self.setGraphicsEffect(self.opacity)

        # set initial size of the slide out menu when program is opened.  not really sure why the
        # height needs to be -57 instead of just minusing the toolbar height directly....
#        self.setGeometry(-250, self.top_toolbar_height, 0, self.mainwindow_height - 57)

        # set color
        self.window_options_top_widget.setStyleSheet('background-color: steelblue;')

        # make top level layout to add to windows_option_top_widget
        self.settings_toplevel_layout = QVBoxLayout()

        # ======================================================================================
        # ======================================================================================

        # top level at top of window
        self.settings_top_label = QLabel("<b>Settings<b>")

        # set font for label
        myfont = QFont()
        myfont.setPointSize(14)
        self.settings_top_label.setFont(myfont)
        self.settings_top_label.setAlignment(Qt.AlignCenter)

        # ======================================================================================
        # ======================================================================================

        # width/height label on launch
        self.expansion_width_height_label = QLabel("<b>Set Expansion Default Size:</b>")
        self.expansion_width_height_label.adjustSize()
        self.expansion_width_height_label.setAlignment(Qt.AlignHCenter)

        # layout to contain width/height widgets
        self.expansion_layout_size = QHBoxLayout()

        # first add stretch to remove , this is called later after widgets are added.... it has to be called once
        # here and then again after widgets are added to make the horizontal widgets stay all grouped together in the
        # center as the window is stretched... i don't know the exact reason you have to call it at the outset
        # and then against afterwards.......
        self.expansion_layout_size.addStretch()

        self.expansion_width_label = QLabel("Width:")
        self.expansion_width_label.adjustSize()

        self.expansion_width_edit = QLineEdit()
        self.expansion_width_edit.setMaximumWidth(50)
        self.expansion_width_edit.setStyleSheet("background-color: lightgrey")
        self.expansion_width_edit.setText(str(self.windowsize_expansion_width))

        self.expansion_height_label = QLabel("Height:")
        self.expansion_height_label.adjustSize()

        self.expansion_height_edit = QLineEdit()
        self.expansion_height_edit.setMaximumWidth(50)
        self.expansion_height_edit.setStyleSheet("background-color: lightgrey")
        self.expansion_height_edit.setText(str(self.windowsize_expansion_height))

        # add widgets to layout
        self.expansion_layout_size.addWidget(self.expansion_width_label)
        self.expansion_layout_size.addWidget(self.expansion_width_edit)
        self.expansion_layout_size.addWidget(self.expansion_height_label)
        self.expansion_layout_size.addWidget(self.expansion_height_edit)

        # add stretch so horizontal items are all next to each other (see note for the 1st time this is called for
        # why this has to be called a second time after widgets are added).
        self.expansion_layout_size.addStretch()

        # ======================================================================================
        # ======================================================================================

        # create checkbox for state for window when it's shrunk
        self.ontop_shrink = QCheckBox("Always On Top When Window Shrunk")
        if self.always_on_top == True:
            self.ontop_shrink.setChecked(True)

        # create checkbox for state for window when it's shrunk
        self.ontop_always = QCheckBox("Always On Top")
        if self.always_on_top_forever == True:
            self.ontop_always.setChecked(True)

        # ======================================================================================
        # ======================================================================================

        # create checboxes to set collapsable on startup
        self.collapse_left_window = QCheckBox("Left Window Collapsed On Startup")
        if self.left_window_collapse == True:
            self.collapse_left_window.setChecked(True)

        self.collapse_top_window = QCheckBox("Top Window Collapsed On Startup")
        if self.top_window_collapse == True:
            self.collapse_top_window.setChecked(True)

        # ======================================================================================
        # ======================================================================================

        # create save and restore buttons for the settings
        self.save_buttons = QHBoxLayout()
        self.save_buttons.addStretch()

        self.save_main_settings = QPushButton("Save Settings")
        self.save_main_settings.adjustSize()
        self.save_main_settings.setStyleSheet("background-color: lightgrey")
        self.save_main_settings.clicked.connect(self.save_personal_settings)

        self.default_main_settings = QPushButton("Restore Defaults")
        self.default_main_settings.adjustSize()
        self.default_main_settings.setStyleSheet("background-color: lightgrey")
        self.default_main_settings.clicked.connect(self.changeto_default_settings)

        self.save_buttons.addWidget(self.save_main_settings)
        self.save_buttons.addSpacing(20)
        self.save_buttons.addWidget(self.default_main_settings)

        self.save_buttons.addStretch()

        # ======================================================================================
        # ======================================================================================

        # add items to the top level settings layout
        self.settings_toplevel_layout.addWidget(self.settings_top_label)
        self.settings_toplevel_layout.addSpacing(100)
        self.settings_toplevel_layout.addWidget(self.expansion_width_height_label)
        self.settings_toplevel_layout.addLayout(self.expansion_layout_size)
        self.settings_toplevel_layout.addSpacing(10)
        self.settings_toplevel_layout.addWidget(self.ontop_shrink)
        self.settings_toplevel_layout.addSpacing(10)
        self.settings_toplevel_layout.addWidget(self.ontop_always)
        self.settings_toplevel_layout.addSpacing(10)
        self.settings_toplevel_layout.addWidget(self.collapse_left_window)
        self.settings_toplevel_layout.addSpacing(10)
        self.settings_toplevel_layout.addWidget(self.collapse_top_window)
        self.settings_toplevel_layout.addSpacing(20)
        self.settings_toplevel_layout.addLayout(self.save_buttons)

        self.settings_toplevel_layout.addStretch()

        # add top level layout to top level widget
        self.window_options_top_widget.setLayout(self.settings_toplevel_layout)

        #set top widget into top scroll
        self.setWidget(self.window_options_top_widget)


    def save_personal_settings(self):

        # make list to put current settings into a list for output to txt later
        self.user_changes = []

        # change global variables based on what is currently in settings
        self.windowsize_expansion_width = int(self.expansion_width_edit.text())
        self.windowsize_expansion_height = int(self.expansion_height_edit.text())

        if self.ontop_shrink.checkState() == 2:
            self.always_on_top = True
        else:
            self.always_on_top = False

        # this is the only setting that needs to be changed immediately when changes are saved (currently)
        if self.ontop_always.checkState() == 2:
            self.always_on_top_forever = True
        else:
            self.always_on_top_forever = False

        if self.collapse_left_window.checkState() == 2:
            self.left_window_collapse = True
        else:
            self.left_window_collapse = False

        if self.collapse_top_window.checkState() == 2:
            self.top_window_collapse = True
        else:
            self.top_window_collapse = False

        # add all setting values to list for output to txt
        self.user_changes.append("windowsize_expansion_width=" + str(self.windowsize_expansion_width))
        self.user_changes.append("windowsize_expansion_height=" + str(self.windowsize_expansion_height))
        self.user_changes.append("always_on_top=" + str(self.always_on_top))
        self.user_changes.append("always_on_top_forever=" + str(self.always_on_top_forever))
        self.user_changes.append("left_window_collapse=" + str(self.left_window_collapse))
        self.user_changes.append("top_window_collapse=" + str(self.top_window_collapse))

        # create new text file and write to it the setting sand values
        # note: user_settings variable for name of txt file is a global variable
        output_file = open(user_path_settings, 'w')
        for i in self.user_changes:
            output_file.write(str(i) + "\n")
        output_file.close()

        # emit to main UI for update to global variables for window state changes
        self.onsettingsChanged.emit(self.windowsize_expansion_width, self.windowsize_expansion_height,
                                    self.always_on_top, self.always_on_top_forever, self.left_window_collapse,
                                    self.top_window_collapse)

    # when default buttons is hit to change the settings back to default
    def changeto_default_settings(self):

        # default settings (same as what's called out at start of program) and resets global variables
        self.windowsize_expansion_width = 910
        self.windowsize_expansion_height = 715
        self.always_on_top = False
        self.always_on_top_forever = False
        self.left_window_collapse = False
        self.top_window_collapse = False

        # change what's listed in settings window UI to the default
        self.expansion_width_edit.setText(str(self.windowsize_expansion_width))
        self.expansion_height_edit.setText(str(self.windowsize_expansion_height))
        self.ontop_shrink.setChecked(self.always_on_top)
        self.ontop_always.setChecked(self.always_on_top_forever)
        self.collapse_left_window.setChecked(self.left_window_collapse)
        self.collapse_top_window.setChecked(self.top_window_collapse)

        # make list for window settings to re-write the main_settings txt file in case in gets changed
        self.default_settings = []

        # add all setting values to list for output to txt
        self.default_settings.append("windowsize_expansion_width=" + str(self.windowsize_expansion_width))
        self.default_settings.append("windowsize_expansion_height=" + str(self.windowsize_expansion_height))
        self.default_settings.append("always_on_top=" + str(self.always_on_top))
        self.default_settings.append("always_on_top_forever=" + str(self.always_on_top_forever))
        self.default_settings.append("left_window_collapse=" + str(self.left_window_collapse))
        self.default_settings.append("top_window_collapse=" + str(self.top_window_collapse))

        # rewrite user text file settings with default settings
        output_file = open(user_path_settings, 'w')
        for i in self.default_settings:
            output_file.write(str(i) + "\n")
        output_file.close()

        # emit to main UI for update to global variables for window state changes
        self.onsettingsChanged.emit(self.windowsize_expansion_width, self.windowsize_expansion_height,
                                    self.always_on_top, self.always_on_top_forever, self.left_window_collapse,
                                    self.top_window_collapse)

if __name__ == "__main__":
    pass
