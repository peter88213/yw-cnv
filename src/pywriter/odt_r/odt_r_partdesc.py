"""Provide a class for ODT invisibly tagged part descriptions import.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.odt_r.odt_r_chapterdesc import OdtRChapterDesc


class OdtRPartDesc(OdtRChapterDesc):
    """ODT part summaries file reader.

    Parts are chapters marked in yWriter as beginning of a new section.
    Import a synopsis with invisibly tagged part descriptions.
    """
    DESCRIPTION = _('Part descriptions')
    SUFFIX = '_parts'
