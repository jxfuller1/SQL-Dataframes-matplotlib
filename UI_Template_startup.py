import getpass
import os

# global variables for window settings options, the settings values are set globally here just in case
# the default settings file doesn't exist or the user settings file doesn't exist for some reason
windowsize_expansion_width = 910
windowsize_expansion_height = 715
always_on_top = False
always_on_top_forever = False
left_window_collapse = False
top_window_collapse = False

# get settings function for interface
def getmain_settings(path):
    # pass globals to here so the global values can be changed
    global windowsize_expansion_width, windowsize_expansion_height, always_on_top, always_on_top_forever, left_window_collapse, top_window_collapse

    main_settings = open(path, "r")
    main_settings_list = main_settings.readlines()

    # pass all settings for txt file to global variables
    for i in main_settings_list:
        if "windowsize_expansion_width=" in i:
            windowsize_expansion_width = int(i.split("=")[1])
        if "windowsize_expansion_height=" in i:
            windowsize_expansion_height = int(i.split("=")[1])
        if "always_on_top" == i:
            # use eval to turn string into a bool
            always_on_top = eval(i.split("=")[1])
        if "left_window_collapse=" in i:
            # use eval to turn string into a bool
            left_window_collapse = eval(i.split("=")[1])
        if "top_window_collapse=" in i:
            # use eval to turn string into a bool
            top_window_collapse = eval(i.split("=")[1])
        if "always_on_top_forever=" in i:
            always_on_top_forever = eval(i.split("=")[1])

def startup_settings():

    # get computer user name, this is used for peoples custom setting initialization files
    user = getpass.getuser()

    # user main settings file name
    # to work for relative path for relative location of the exe/python exe
    user_settings = "\\main_settings_" + user + ".txt"

    # get current working directory (needed if using on different computers
    my_path = os.path.abspath(os.path.dirname(__file__))

    # this would be the user settings path relative to exe location
    user_path_settings = my_path + user_settings

    # directory for default settings using relative path settings
    default_path_settings = my_path + "\\main_settings.txt"

    # if path exists get settings for the interface, if default settings file nor user settings
    # file doesn't exist for whatever reason use the global settings values
    if os.path.exists(user_path_settings):
        getmain_settings(user_path_settings)
    else:
        if os.path.exists(default_path_settings):
            getmain_settings(default_path_settings)

    return windowsize_expansion_width, windowsize_expansion_height, always_on_top, always_on_top_forever, \
           left_window_collapse, top_window_collapse
