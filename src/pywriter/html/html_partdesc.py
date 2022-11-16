"""Provide a class for html invisibly tagged part descriptions import.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.html.html_chapterdesc import HtmlChapterDesc


class HtmlPartDesc(HtmlChapterDesc):
    """HTML part summaries file representation.

    Parts are chapters marked in yWriter as beginning of a new section.
    Import a synopsis with invisibly tagged part descriptions.
    """
    DESCRIPTION = _('Part descriptions')
    SUFFIX = '_parts'
