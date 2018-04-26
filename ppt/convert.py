
import os

from pptx import Presentation

from ppt.PATH_dou import SOURCE, DEST
from ppt.to_word import convert, convert_ppt

OVERWRITE = True

files = filter(lambda f: os.path.isfile(SOURCE + os.sep + f), os.listdir(SOURCE))
for file in files:
    if not OVERWRITE:
        if os.path.exists(DEST + os.sep + os.path.splitext(file)[0] + '.docx'):
            print('File: %s \t already converted' % file)
            continue
    ext_name = os.path.splitext(file)[-1]
    if ext_name == '.pptx':
        print('Converting', file)
        prs = Presentation(SOURCE + os.sep + file)
        convert(prs, DEST + os.sep + os.path.splitext(file)[0] + '.docx')
    elif ext_name == '.ppt':
        print('Converting', file)
        convert_ppt(SOURCE + os.sep + file, DEST + os.sep + os.path.splitext(file)[0] + '.docx')

