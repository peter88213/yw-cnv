"""Provide a factory class for a document object to read and a new yWriter project.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
from pywriter.pywriter_globals import *
from pywriter.converter.file_factory import FileFactory
from pywriter.yw.yw7_file import Yw7File
from pywriter.html.html_import import HtmlImport
from pywriter.html.html_outline import HtmlOutline
from pywriter.html.html_fop import read_html_file


class NewProjectFactory(FileFactory):
    """A factory class that instantiates a document object to read, 
    and a new yWriter project.

    Public methods:
        make_file_objects(self, sourcePath, **kwargs) -- return conversion objects.

    Class constant:
        DO_NOT_IMPORT -- list of suffixes from file classes not meant to be imported.    
    """
    DO_NOT_IMPORT = ['_xref', '_brf_synopsis']

    def make_file_objects(self, sourcePath, **kwargs):
        """Instantiate a source and a target object for creation of a new yWriter project.

        Positional arguments:
            sourcePath -- str: path to the source file to convert.

        Return a tuple with two elements:
        - sourceFile: a Novel subclass instance
        - targetFile: a Novel subclass instance
        
        Raise the "Error" exception in case of error. 
        """
        if not self._canImport(sourcePath):
            raise Error(f'{_("This document is not meant to be written back")}.')

        fileName, __ = os.path.splitext(sourcePath)
        targetFile = Yw7File(f'{fileName}{Yw7File.EXTENSION}', **kwargs)
        if sourcePath.endswith('.html'):
            # The source file might be an outline or a "work in progress".
            content = read_html_file(sourcePath)
            if "<h3" in content.lower():
                sourceFile = HtmlOutline(sourcePath, **kwargs)
            else:
                sourceFile = HtmlImport(sourcePath, **kwargs)
            return sourceFile, targetFile

        else:
            for fileClass in self._fileClasses:
                if fileClass.SUFFIX is not None:
                    if sourcePath.endswith(f'{fileClass.SUFFIX}{fileClass.EXTENSION}'):
                        sourceFile = fileClass(sourcePath, **kwargs)
                        return sourceFile, targetFile

            raise Error(f'{_("File type is not supported")}: "{norm_path(sourcePath)}".')

    def _canImport(self, sourcePath):
        """Check whether the source file can be imported to yWriter.
        
        Positional arguments: 
            sourcePath -- str: path of the file to be ckecked.
        
        Return True, if the file located at sourcepath is of an importable type.
        Otherwise, return False.
        """
        fileName, __ = os.path.splitext(sourcePath)
        for suffix in self.DO_NOT_IMPORT:
            if fileName.endswith(suffix):
                return False

        return True
