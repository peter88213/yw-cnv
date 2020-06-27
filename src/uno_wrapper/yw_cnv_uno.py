"""Import and export yWriter data. 

Standalone yWriter file converter with basic error handling 

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY

from pywriter.converter.yw_cnv import YwCnv
from uno_wrapper.uno_tools import *


class YwCnvUno(YwCnv):
    """Converter for yWriter project files.
    Variant with UNO UI.
    """

    def confirm_overwrite(self, filePath):
        result = msgbox('Overwrite existing file "' + filePath + '"?',
                        'WARNING', BUTTONS_YES_NO, 'warningbox')

        if result == YES:
            return True

        else:
            return False
