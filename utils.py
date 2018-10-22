import mimetypes
import os
import zipfile

import pandas as pd
from pandas import ExcelFile


def walkdir(path):
    """
    Recursively yield DirEntry objects for given directory.
    :param path: the root directory to start processing from
    """
    print('Proccessing directory :{0}'.format(path))
    for entry in os.scandir(path):
        if entry.is_file():
            yield entry
        elif entry.is_dir(follow_symlinks=False):
            yield from walkdir(entry.path)

def dispatch(path, mtype):

    try:
        if 'zip' in mtype:
            zf = zipfile.ZipFile(path)
            for zinfo in zf.infolist():
                if zinfo.file_size > 0:
                    mtype = mimetypes.guess_type(zinfo.filename)[0]
                    dispatch(zinfo.filename, mtype)
        print('{0} is a {1}'.format(path, mtype))
        if 'spreadsheet' in mtype or 'ms-excel' in mtype:
            openExcelFile(path)
    except:
        print('File : {0} has a mime type that is not supported !!!'.format(path))

    # print filename and file type

def openExcelFile(data):
    '''
        Read and/or process MS Excel file
    :param data: excel file data or path to file to open it
    :return:
    '''
    if type(data) is str:
        df = pd.read_excel(data)
    else:
        df = ExcelFile(data)

    print(df.columns)