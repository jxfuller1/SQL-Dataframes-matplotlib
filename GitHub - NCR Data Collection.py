import sys
import os
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QSplashScreen, QApplication

# get current working directory (needed if using on different computers
my_path = os.path.abspath(os.path.dirname(__file__))

# splash screen
app = QApplication(sys.argv)
splash = QSplashScreen(QPixmap(my_path + "\\ncr_graph_splash.png"))
splash.show()

from PyQt5.QtCore import Qt, QSize, QPropertyAnimation, QRect, QTimer, QTime
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtWidgets import QApplication, QMainWindow, QHBoxLayout, QWidget, QVBoxLayout, QStackedLayout, \
    QStackedWidget, QScrollArea, QSizePolicy, QSplitter, QToolBar, QAction, QToolButton, QStatusBar, QActionGroup, \
    QMenu, QGroupBox, QLabel
import getpass

# UI portions of each specific program offloaded to separate file
import Github_Data_Program
# get startup settings for UI
import UI_Template_startup
# top buttons in UI
import UI_Template_top_buttons
# left buttons in UI
import UI_Template_leftside_buttons
# setup for window settings window in UI
import UI_Template_Window_settings

import win32print
import win32com.client

# get printers for UI (local printers)
printer_names = ([printer[2] for printer in win32print.EnumPrinters(2)])
# network printers
printer_names1 = ([printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_CONNECTIONS, None, 1)])
# combine local and network printers
printer_names = printer_names + printer_names1


# NOTE: when adding more settings options for UI, you WILL need to:
# ------- add UI element to UI_Template_Window_settings and account for it in all places there
# ------- account for it in UI_Template_startup
# ------- add variable for it in QMainWindow
# ------- add the new var for the self.windows_options callout and update_window_and_settings function

# Note: when adding a NEW program
# ----- create file for the program with how you want the UI to look for each page for the UI
#       **see UI_template_program_generic
# ----- add import statement to Qstackedlayout file (MUST be a class with Qstackedlayout)
# ----- add buttons you want program to have to program_buttons list
# ----- add program name to total_programs list
# ----- initialize the file in the 'init" section of Actions() and add it to self.programs_list

# note: if you want something to activate when the program leftside buttons in UI are changed
        # refer to switch_leftsidebuttons_UI function


# enter number of program names you want to appear in UI, put into this list
total_programs = ["NCR Data"]
# enter number of buttons for each program and their names that you want to appear on left side list in UI
program_buttons = [["Home", "Charts", "Add Months Data", "JOB Filters", "NCR Filters", "VENDOR Filters", "About"]]

# this is for scaling with higher res monitors
if hasattr(Qt, 'AA_EnableHighDpiScaling'):
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

# resource_path function MAY be needed instead of using my_path
# when turning file into an exe for icons/images/etc... not sure if it'll be needed yet
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class Actions(QMainWindow):

    def __init__(self):
        super().__init__()

        # get startup settings
        self.windowsize_expansion_width, \
            self.windowsize_expansion_height, \
            self.always_on_top, \
            self.always_on_top_forever, \
            self.left_window_collapse, \
            self.top_window_collapse = UI_Template_startup.startup_settings()

        # initialize the layouts for each program, which will be added to the UI in mainwork_UI_Layout function
        # initializing them here for ease of use when adding programs
        self.ncrdata_mainwork_layout = Github_Data_Program.Program_layout(self)
        self.ncrdata_mainwork_layout.resize_window.connect(self.ongraphChange)
        self.ncrdata_mainwork_layout.print_selection.connect(self.check_print_menu)

        # put all programs into a list for the mainwork_ui_layout function
        self.program_list = [self.ncrdata_mainwork_layout]

        # default window size and start location on startup and default shrinksize
        self.xpos = 400
        self.ypos = 200
        self.mainwidthsize = 910
        self.mainheightsize = 715
        self.shrinksize_width = 250
        self.shrinksize_height = 0
        # tool bar height is 57 depending on size of icons... don't know why framegeometry on the toolbar
        # plus statusbar height, need this height for the logic of shrinking/expanding window
        self.default_toolbar_status_height = 77

        self.initUI()

    # signal from ncr data program for resizing the window just slightly so that the graphs update properly...
    # couldn't figure out a way to do it anyother way
    def ongraphChange(self, value):
        x, y, width, height = self.geometry().getRect()

        adjusted_width = width+1
        adjusted_height = height+1

        self.setGeometry(x, y, adjusted_width, adjusted_height)
        self.setGeometry(x, y, width, height)

    # check print menu to see which printer is selected
    def check_print_menu(self, value):
        # initialize printer value
        printer = "None"

        # get which printer is checkmarked, if none send message printer not chosen
        for i in self.printer_action_group.actions():
            if i.isChecked():
                printer = i.text()

        self.ncrdata_mainwork_layout.print_charts(printer)

    # resize event for main window for changing the size of the slide out settings menu
    def resizeEvent(self, event):
        # resize slideout menu if maiwindow is resized.  not really sure why the
        # height needs to be -57 instead of just minusing the height of the toolbar itself
        self.windows_options.resize(self.windows_options.frameGeometry().width(),
                                   self.frameGeometry().height() - self.default_toolbar_status_height)

        # changes shrink/expand icon based on window height to match when you can expand or shrink window
        if self.frameGeometry().height() <= self.default_toolbar_status_height:
            self.shrink.setIcon(QIcon(my_path + "\\expand-multimedia-option-svgrepo-com"))
            self.shrink.setToolTip("Expand")
        else:
            self.shrink.setIcon(QIcon(my_path + "\\shrink-screen-filled-svgrepo-com"))
            self.shrink.setToolTip("Shrink")

    def initUI(self):

        # =========================================================================================
        # ====================set window settings==================================================
        # =========================================================================================

        # set size of window at startup, this will be called from ini file later, not coded  in yet
        self.setGeometry(self.xpos, self.ypos, self.mainwidthsize, self.mainheightsize)

        # set windows flags for minimize/close/maximize buttons based on init txt file
        if self.always_on_top_forever == True:
            self.setWindowFlags(Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint
                                | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)

        self.setWindowTitle("NCR Data Collection")

        # =========================================================================================
        # ===================above are window settings=============================================
        # =========================================================================================

        # =========================================================================================
        # ===================create toolbar settings===============================================
        # =========================================================================================

        # add toolbar for options/window expansion/contraction
        self.top_toolbar = QToolBar("Options", self)
        self.addToolBar(Qt.TopToolBarArea, self.top_toolbar)

        # prevent toolbar from being movable to other widges of window or on it's own window
        self.top_toolbar.setMovable(False)

        # prevent right click options of toolbar
        self.top_toolbar.setContextMenuPolicy(Qt.PreventContextMenu)

        # add items to toolbar
        self.openmenu = QAction(QIcon(my_path + "\\options-lines-svgrepo-com"), "Settings", self)
        self.openmenu.triggered.connect(self.windowselect_options)
        self.top_toolbar.addAction(self.openmenu)

        # using this method for adding icon to toolbar
        self.shrink = QToolButton()
        self.shrink.setIcon(QIcon(my_path + "\\shrink-screen-filled-svgrepo-com"))
        self.shrink.setToolTip("Shrink")
        self.shrink.clicked.connect(self.change_mainwindow_size)
        self.top_toolbar.addWidget(self.shrink)


        # ========================= create printer menu option in tool bar ========================
        # ========================= create printer menu option in tool bar ========================
        # ========================= create printer menu option in tool bar ========================
        self.print_menu = QMenu()

        # put printer menu items in actiongroup so that only 1 can be checked at a time
        self.printer_action_group = QActionGroup(self.print_menu)

        # add menu selection for every printer so user can choose
        for i in printer_names:
            menu_item = QAction(i, self.print_menu)
            menu_item.setCheckable(True)

            # add menu item to menu and action group
            self.print_menu.addAction(menu_item)
            self.printer_action_group.addAction(menu_item)

        self.printer_action_group.setExclusive(True)

        self.print_toolbutton = QToolButton()
        self.print_toolbutton.setIcon(QIcon(my_path + "\\print-svgrepo-com"))
        self.print_toolbutton.setMenu(self.print_menu)
        self.print_toolbutton.setPopupMode(QToolButton.InstantPopup)

        self.top_toolbar.addWidget(self.print_toolbutton)
        # ========================= create printer menu option in tool bar ========================
        # ========================= create printer menu option in tool bar ========================
        # ========================= create printer menu option in tool bar ========================


        # set icons size for toolbar
        self.top_toolbar.setIconSize(QSize(15, 15))

        # =========================================================================================
        # ===================above toolbar creation================================================
        # =========================================================================================

        # =========================================================================================
        # ===================create top level widgets/layouts for UI===============================
        # =========================================================================================

        # set top level scroll area for top level widget to go into, gets added to central widget for
        # main window
        self.toplevel_scroll = QScrollArea()

        # add splitter for top buttons at top, gets added to toplevel_scroll
        self.toplevel_splitter = QSplitter()
        self.toplevel_splitter.setOrientation(Qt.Vertical)

        # set size policy so that the mainwindow can be scrunched as small as possible
        self.toplevel_scroll.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)

        # set resizable so widgets don't take up small areas in layouts, set scrollarea widget
        self.toplevel_scroll.setWidgetResizable(True)
        self.toplevel_scroll.setWidget(self.toplevel_splitter)

        # top level layout that all layouts/widgets get added to, this gets added to toplevel_widget
        self.toplevel_layout = QVBoxLayout()

        # =========================================================================================
        # ===================above is top level widgets for UI=====================================
        # =========================================================================================

        # =========================================================================================
        # ====================secondary level layouts for UI=======================================
        # =========================================================================================

        # make top level a splitter so that user can move the window the way they want
        self.secondarylevel_splitter = QSplitter()

        # secondary top level layout, add horizontal layout to the vertical layout
        self.secondarylevel_layout = QHBoxLayout()
        # set zero spacing between items added to the qhboxlayout
        self.secondarylevel_layout.setSpacing(0)

        # =========================================================================================
        # ====================above secondary level layouts for UI=================================
        # =========================================================================================

        # =========================================================================================
        # ====================create top level buttons of UI=======================================
        # =========================================================================================

        # create top buttons for UI
        # create main widget for top buttons
        self.topbuttons_widget = QWidget()

        # set object name for style sheet in order to just make the widget grey and not the buttons too
        self.topbuttons_widget.setObjectName("button_widget")
        self.topbuttons_widget.setStyleSheet("QWidget#button_widget {background-color: gray}")

        # call file that populates top buttons for UI
        self.top_button_layout = UI_Template_top_buttons.main_top_buttons(self, total_programs)
        self.top_button_layout.ontopbuttonChanged.connect(self.switch_leftsidebuttons_UI)

        # set layout to main UI Widget
        self.topbuttons_widget.setLayout(self.top_button_layout)

        # set height limit because of Qsplitter that allows user to expand window height for this widget
        # i don't really want it to expand more, just less
        self.topbuttons_widget.setMaximumHeight(45)

        # =========================================================================================
        # ====================above is top level buttons of UI=====================================
        # =========================================================================================

        # =========================================================================================
        # ====================create leftside buttons of UI========================================
        # =========================================================================================

        # create stacked layout, when top_buttons of UI are clicked, the stacked layout
        # will change indexes and change the left side buttons on the U
        self.left_buttons = UI_Template_leftside_buttons.main_leftside_buttons(self, program_buttons)
        # connectoin for when button is clicked for leftside
        self.left_buttons.programpageChanged.connect(self.switch_mainwork_layout_UI)

        # =========================================================================================
        # ====================create leftside buttons of UI=================================
        # =========================================================================================

        # =========================================================================================
        # ====================create create mainwork pages of UI for each program==================
        # =========================================================================================

        # program main work stackedwidget.  When left side buttons are clicked, the stackedwidget
        # will change indexes and change the UI where the main work is done

        self.upper_mainwork_layout_UI = QStackedWidget()

        # add each program from self.program_list to UI
        for i in self.program_list:
            # upper widget that gets added to QStackedwidget
            program_widget = QWidget()

            # set layout of widget
            program_widget.setLayout(i)

            # add widget that has stackedlayout to Qstackedwidget
            self.upper_mainwork_layout_UI.addWidget(program_widget)

        # =========================================================================================
        # ====================above is create mainwork pages of UI for each program================
        # =========================================================================================

        # =========================================================================================
        # ====================set all the widgets/layouts together for order creation of UI========
        # =========================================================================================

        # set top level buttons of UI to main layout
        self.toplevel_layout.addWidget(self.topbuttons_widget)

        # add tabs and left side buttons to secondary layout on UI that is below the top level buttons
        self.secondarylevel_layout.addWidget(self.left_buttons)

        # pass 2 into here so that only the mainwork_layout stretches (not really needed to set the 2 in here
        # as this goes into the splitter, where it needs to be set, but just keeping on here
        # if I want to get rid of the splittler later
        self.secondarylevel_layout.addWidget(self.upper_mainwork_layout_UI, 2)

        # set layout of splitter and set stretch factor of the mainwork section of the UI so that
        # only that stretches, this will then be added to the top level layout
        self.secondarylevel_splitter.setLayout(self.secondarylevel_layout)
        self.secondarylevel_splitter.setStretchFactor(1, 2)

        # add secondary level layout to the top level layout
        self.toplevel_layout.addWidget(self.secondarylevel_splitter)

        # set layout of toplevel to splitter and set stretch factor so only mainwork UI window stretches
        self.toplevel_splitter.setLayout(self.toplevel_layout)
        self.toplevel_splitter.setStretchFactor(1, 2)

        # SET UI SETTINGS BASED ON INIT FILE AT STARTUP
        # if collapsed on startup settings is selected, collapse top window
        if self.top_window_collapse == True:
            # sets top window to being collapsed on startup... i don't actually understand the parameters
            # of [-1, 1], but it works......
            self.toplevel_splitter.setSizes([-1, 1])
        # collapse left side window of UI
        if self.left_window_collapse == True:
            self.secondarylevel_splitter.setSizes([-1, 1])

        # remove space borders between mainwindow and layout
        self.toplevel_layout.setContentsMargins(0, 0, 0, 0)

        # remove spacing between layouts (upper and lower part of UI)
        self.toplevel_layout.setSpacing(0)

        # set toplevel_scroll to main window
        self.setCentralWidget(self.toplevel_scroll)

        # =========================================================================================
        # ====================set all the widgets/layouts together for order creation of UI========
        # =========================================================================================

        # ================================================================================================
        # =============================create windows settings slide out of UI============================
        # ================================================================================================

        # pass window settings to the settings window so that settings in UI can be filled out correctly
        self.windows_options = UI_Template_Window_settings.window_settings(self, self.windowsize_expansion_width,
                                                                           self.windowsize_expansion_height, self.always_on_top,
                                                                           self.always_on_top_forever, self.left_window_collapse,
                                                                           self.top_window_collapse)

        self.windows_options.onsettingsChanged.connect(self.update_window_and_settings)

        self.windows_options.setGeometry(-250, self.top_toolbar.frameGeometry().height(), 0,
                                         self.frameGeometry().height() - self.default_toolbar_status_height)

        self.windows_options.hide()

        # ================================================================================================
        # ==========================above creates setting slide out window of UI==========================
        # ================================================================================================

        # ================================================================================================
        # ==========================create status bar=====================================================
        # ================================================================================================
        # use status bar for program updates
        self.statusbar = QStatusBar()

        user = getpass.getuser()
        # status bar groupbox labels
        status_bar_labels = [user, "", "", "", "Version: 2.0.0"]

        # keep references of the labels in the status bar for updating as necessary
        self.status_label_references = []

        myfont = QFont()
        myfont.setPointSize(8)

        # create groupboxes with labels for statusbar and add them
        for i in status_bar_labels:
            group_box = QGroupBox()
            group_box.setFixedHeight(18)

            layout = QVBoxLayout()
            layout.setContentsMargins(0, 0, 0, 0)

            label = QLabel(i)
            label.setFont(myfont)

            layout.addWidget(label)
            group_box.setLayout(layout)
            self.statusbar.addWidget(group_box, stretch=1)

            self.status_label_references.append(label)

        # Create a timer to update the clock every second
        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self.update_time)
        self.clock_timer.start(1000)

        self.setStatusBar(self.statusbar)
        # ================================================================================================
        # ==========================create status bar=====================================================
        # ================================================================================================

        self.show()

    # update clock and pass ot statusbar label
    def update_time(self):
        current_time = QTime.currentTime()
        formatted_time = current_time.toString('hh:mm:ss')
        self.status_label_references[2].setText(formatted_time)

    # update global values for settings based on inputs from settings part of UI, and change mainwindow flags
    def update_window_and_settings(self, window_exp_width, window_exp_height, on_top, on_top_forever, left_collapse, top_collapse):

        # change global values for settings
        self.windowsize_expansion_width = window_exp_width
        self.windowsize_expansion_height = window_exp_height
        self.always_on_top = on_top
        self.always_on_top_forever = on_top_forever
        self.left_window_collapse = left_collapse
        self.top_window_collapse = top_collapse

        # set window flag
        if self.always_on_top_forever == True:
            # add always on top (in a try statement as im too lazy to make a check if the windowflag is already there)
            try:
                self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
                # reload window (has to be done after changing windows flags
                self.show()
            except:
                pass
        else:
            try:
                # remove always on top
                self.setWindowFlags(self.windowFlags() & ~Qt.WindowStaysOnTopHint)
                # reload window (has to be done after changing windows flags
                self.show()
            except:
                pass

    # changes window size for the expansion/shrink toolbar button
    def change_mainwindow_size(self):
        # get current x/y position
        self.current_x_pos = self.frameGeometry().x()
        self.current_y_pos = self.frameGeometry().y()

        # create animation property
        self.mainwindow_animation = QPropertyAnimation(self, b'size')

        # animation speed
        self.mainwindow_animation.setDuration(250)

        # if current state of window is above shrinkwidth or above height width, make window smaller
        # NOTE: Height check is checking if greater than 57 due to height of toolbar
        # for some reason doing self.framegeometry on the tool bar returns 25... instead of 57....
        # i don't know why
        if self.frameGeometry().width() > self.shrinksize_width or self.frameGeometry().height() > self.default_toolbar_status_height:
            self.mainwindow_animation.setEndValue(QSize(self.shrinksize_width, self.shrinksize_height))
            self.mainwindow_animation.start()

            # if always on top setting is set, make shrunk window on top
            if self.always_on_top == True:
                self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
                self.show()

        # if current state of window is minimized, then make larger
        if self.frameGeometry().width() == self.shrinksize_width or self.frameGeometry().height() == self.default_toolbar_status_height:
            # expand window to size set in settings
            self.mainwindow_animation.setEndValue(QSize(self.windowsize_expansion_width, self.windowsize_expansion_height))
            self.mainwindow_animation.start()

            # remove on top flag, but only if the setting is set to False in the settings
            if self.always_on_top_forever == False:
                self.setWindowFlags(self.windowFlags() & ~Qt.WindowStaysOnTopHint)
            self.show()

    # when options toolbar button hit
    def windowselect_options(self):
        # when slide out menu button hit, show and raise it to the top
        self.windows_options.show()
        self.windows_options.raise_()

        # create animation for the slide out window
        self.window_animation = QPropertyAnimation(self.windows_options, b'geometry')

        # get current state of window
        self.window_animation.setStartValue(self.windows_options.frameGeometry())

        # slide out animation speed
        self.window_animation.setDuration(250)

        # if current state of window is minimized, then make bigger
        if self.windows_options.frameGeometry().x() <= -250:
            self.window_animation.setEndValue(
                QRect(0, self.top_toolbar.frameGeometry().height(), 250, self.frameGeometry().height() - self.default_toolbar_status_height))
            self.window_animation.start()

        # if current state of window is maximized, then make smaller
        if self.windows_options.frameGeometry().x() >= 0:
            self.window_animation.setEndValue(
                QRect(-250, self.top_toolbar.frameGeometry().height(), 0, self.frameGeometry().height() - self.default_toolbar_status_height))
            self.window_animation.start()

    # switch leftside buttons and mainwork layout part of UI when top program buttons are hit in UI
    def switch_leftsidebuttons_UI(self, sender):

        # change left side layout buttons and mainwork window depending on which button is clicked on in the top of the UI
        self.left_buttons.leftside_buttons_layout.setCurrentIndex(self.top_button_layout.indexOf(sender))
        self.upper_mainwork_layout_UI.setCurrentIndex(self.top_button_layout.indexOf(sender))

    # set index of page to be displayed based on which program page button was hit, values are emitted from each program
    # in the stackedlayout for each program
    def switch_mainwork_layout_UI(self, sub_index):

        # get index of the stackedlayout for the leftside_buttons of UI in order to set the correct index
        # of the stackedlayout mainwork section of the UI for the corresponding program
        index = self.left_buttons.parentWidget().findChild(QStackedLayout).currentIndex()

        # set corresponding program page (stackedlayout) for button that was clicked from the self.program_list
        self.program_list[index].setCurrentIndex(sub_index)

    # when window is closed, close the sql connection
  #  def closeEvent(self, event):
  #      self.ncrdata_mainwork_layout.close()

if __name__ == "__main__":
   # app = QApplication(sys.argv)

    window = Actions()
    splash.finish(window)
    sys.exit(app.exec_())
