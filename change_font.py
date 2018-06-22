#!/usr/bin/python3
'''Change workbook font to Tohama to solve the problem of Tableau book font in Tableau version 10.4-10.5

Usage
=====
python3 change_font.py <twb or twbx file> <output directory>

Note
====
The code directly insert the style element after </preferences> tag.
It doesn't check anything including whether the style has already been inserted.
This code is to quickly add font to many Tableau workbook to save time.
Will need a lot more checks to use in general.

Copyright 2018 - Prapat Suriyaphol

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

'''
import os
import sys
from pathlib import Path
import zipfile
import shutil

TAHOMA_FONT_STYLE = (
'''  <style>
    <style-rule element='all'>
      <format attr='font-family' value='Tahoma' />
    </style-rule>
  </style>
''')

def add_font(twb_filepath, output_filepath):
    """add font style into twb and save into output_dir"""

    with open(twb_filepath, encoding='utf-8') as twb_file, \
       open(output_filepath,'w',encoding='utf-8') as outf:
      for line in twb_file:
        outf.write(line)
        # if last line is </preferences>, insert tahomo style after that
        if line.strip() == '</preferences>':
          outf.write(TAHOMA_FONT_STYLE)

def get_filename(filepath):
  """return only filename with extension. no path"""
  return Path(filepath).name

def process_twb(input_file, output_dir):
  twb_filename = get_filename(input_file)
  # Create a temporary filename for the new file with added font
  output_filepath = Path(output_dir)/'temp_with_added_font.twb'
  add_font(str(input_file), str(output_filepath))
  # replace file with the correct name
  os.replace(str(output_filepath), str(Path(output_dir)/twb_filename))

def extract_tbwx(input_file, output_dir):
  zip_ref = zipfile.ZipFile(input_file, 'r')
  zip_ref.extractall(str(output_dir))
  zip_ref.close()

def prepare_temp_dir(output_dir):
  '''Prepare temp directory and return the temp directory path'''
  temp_dir = Path(output_dir)/'tempdir'
  # remove the directory first if exist
  if temp_dir.exists():
    shutil.rmtree(str(temp_dir.absolute()))
  temp_dir.mkdir()
  return temp_dir

def get_twb_filepath(temp_dir):
  '''Return full path of twb file'''
  all_twb_files = list(temp_dir.glob('*.twb'))
  # make sure that there is only 1 twb file
  assert len(all_twb_files) == 1
  twb_filepath = all_twb_files[0]
  return twb_filepath

def zip_back_to_twbx(temp_dir, twbx_filename):
  '''Compress directory back to the original twbx file'''
  current_dir = os.getcwd()
  os.chdir(temp_dir)
  zipfile = shutil.make_archive('../temp', 'zip')
  os.replace(zipfile, '../'+twbx_filename)
  os.chdir(current_dir)

if __name__ == '__main__':
  # First argument is input filename
  # second argument is output directory
  input_file = sys.argv[1]
  output_dir = sys.argv[2]
  ext = Path(input_file).suffix

  if ext == '.twbx':
    twbx_filename = get_filename(input_file)
    temp_dir = prepare_temp_dir(output_dir)

    # extract tbwx file
    extract_tbwx(input_file, temp_dir)

    # get the twb file inside the extracted directory
    twb_filepath = get_twb_filepath(temp_dir)

    # add tahoma font
    process_twb(twb_filepath, temp_dir)

    # zip file back and remove the temp directory
    zip_back_to_twbx(str(temp_dir), twbx_filename)
    if temp_dir.exists():
      shutil.rmtree(str(temp_dir))
  elif ext == '.twb':
    process_twb(input_file, output_dir)
  else:
    print('Unknown file type')

  print('Done')
