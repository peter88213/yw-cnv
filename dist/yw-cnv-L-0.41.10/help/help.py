"""Show html help files. 

This is a workaround for the apparently not working help interface of LibreOffice.

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""

import os


def show_help():
    scriptLocation = os.path.dirname(__file__)
    os.startfile(scriptLocation + '/help.html')


def show_adv_help():
    scriptLocation = os.path.dirname(__file__)
    os.startfile(scriptLocation + '/help_adv.html')
