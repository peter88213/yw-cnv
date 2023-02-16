"""Provide a class for ODT invisibly tagged "Todo" chapters import.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.odt_r.odt_r_manuscript import OdtRManuscript


class OdtRTodo(OdtRManuscript):
    """ODT "Todo" chapters file reader.

    Import a manuscript with invisibly tagged chapters and scenes.
    """
    DESCRIPTION = _('Todo chapters')
    SUFFIX = '_todo'

    _TYPE = 2
