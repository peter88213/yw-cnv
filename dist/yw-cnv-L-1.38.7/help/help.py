"""Show html help files. 

This is a workaround for the apparently not working help interface of LibreOffice.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import webbrowser
import locale
import uno


def open_helppage(name):
    """Open a localized help page, if any.
    
    Otherwise open the standard help page.
    """
    try:
        lang = locale.getlocale()[0][:2]
    except:
        # Fallback for old Windows versions.
        lang = locale.getdefaultlocale()[0][:2]
    scriptLocation = os.path.dirname(__file__)
    helpFile = f'{scriptLocation}/{name}-{lang}.html'
    if not os.path.isfile(uno.fileUrlToSystemPath(helpFile)):
        helpFile = f'{scriptLocation}/{name}.html'
    webbrowser.open(helpFile)


def show_help():
    open_helppage('help')


def show_adv_help():
    open_helppage('help-adv')
