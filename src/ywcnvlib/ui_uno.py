"""Provide a UNO user interface facade class.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/yw-cnv
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from pywriter.ui.ui import Ui
from ywcnvlib.uno_tools import *


class UiUno(Ui):
    """UI subclass implementing a LibreOffice UNO facade."""

    def ask_yes_no(self, text):
        result = msgbox(text, buttons=BUTTONS_YES_NO, type_msg=WARNINGBOX)
        return result == YES

    def set_info_how(self, message):
        """How's the converter doing?"""
        self.infoHowText = message
        if message.startswith('!'):
            message = message.split('!', maxsplit=1)[1].strip()
            msgbox(message, type_msg=ERRORBOX)
        else:
            msgbox(message, type_msg=INFOBOX)

    def show_warning(self, message):
        """Display a warning message box."""
        msgbox(message, buttons=BUTTONS_OK, type_msg=WARNINGBOX)

