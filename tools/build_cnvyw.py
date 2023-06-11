"""Build a Python script for the LibreOffice "convert yWriter" script.
        
Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os
import sys
# sys.path.insert(0, f'{os.getcwd()}/../../PyWriter/src')
import inliner

SRC = '../src/'
BUILD = '../test/'
SOURCE_FILE = f'{SRC}cnvyw_.py'
TARGET_FILE = f'{BUILD}cnvyw.py'

SCRIPT_CODE = """LOCALE_PATH = f'{os.path.dirname(sys.argv[0])}/locale/'"""
UNO_CODE = """oPackageInfoProvider = CTX.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
sPackageLocation = oPackageInfoProvider.getPackageLocation("org.peter88213.yw-cnv")
packagePath = uno.fileUrlToSystemPath(sPackageLocation)
LOCALE_PATH = f'{packagePath}/yw-cnv/locale/'"""


def main():
    inliner.run(SOURCE_FILE, TARGET_FILE, 'ywcnvlib', '../src/', copyPyWriter=True)
    inliner.run(TARGET_FILE, TARGET_FILE, 'pywriter', '../../PyWriter/src/', copyPyWriter=True)
    # inliner.run(SOURCE_FILE, TARGET_FILE, 'ywcnvlib', '../src/')
    # inliner.run(TARGET_FILE, TARGET_FILE, 'pywriter', '../src/')

    # This is a hack to get the script's location within the UNO context:
    with open(TARGET_FILE, 'r') as f:
        text = f.read()
        text = text.replace(SCRIPT_CODE, UNO_CODE)
    with open(TARGET_FILE, 'w') as f:
        f.write(text)
    print('Done.')


if __name__ == '__main__':
    main()
