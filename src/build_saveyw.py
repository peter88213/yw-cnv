""" Build python script for the LibreOffice "save to yWriter" script.
        
In order to distribute single scripts without dependencies, 
this script "inlines" all modules imported from the pywriter package.

For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import inliner

SRC = '../src/'
BUILD = '../test/'


def main():
    os.chdir(SRC)
    inliner.run('saveyw_.py', BUILD + 'saveyw.py', 'pywriter')
    print('Done.')


if __name__ == '__main__':
    main()
