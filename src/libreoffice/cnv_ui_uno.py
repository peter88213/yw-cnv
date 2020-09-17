"""User interface for the converter: UNO facade

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os

from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX

from pywriter.converter.yw_cnv import YwCnv
from pywriter.converter.file_factory import FileFactory
from pywriter.yw.yw7_tree_creator import Yw7TreeCreator
from libreoffice.uno_tools import *


class CnvUiUno():
    """UI subclass implementing a LibreOffice UNO facade."""

    def ask_yes_no(self, text):
        result = msgbox(text, 'WARNING', buttons=BUTTONS_YES_NO,
                        type_msg=WARNINGBOX)

        if result == YES:
            return True

        else:
            return False

    def set_info_how(self, message):
        """How's the converter doing?"""
        self.infoHowText = message

        if message.startswith('SUCCESS'):
            msgbox(message, type_msg=INFOBOX)

        else:
            msgbox(message, type_msg=ERRORBOX)
