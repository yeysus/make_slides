# Create a set of PowerPoint slides based on Screenshots from Websites and
# Info text in an Excel file.

# sudo pip3 install Pillow
# sudo pip3 install lxml
# sudo pip3 install XlsxWriter
# sudo pip3 install python-pptx

# How to place pictures in a slide:
# https://stackoverflow.com/questions/44275443/python-inserts-pictures-to-powerpoint-how-to-set-the-width-and-height-of-the-pi

# How to calculate aspect ratio:
# https://www.omnicalculator.com/other/aspect-ratio

import sys
# For sleep() in some tests I did.
import time
import pandas as pd
import os, sys, subprocess

from pptx import Presentation
from pptx.dml.color import ColorFormat, RGBColor
from pptx.util import Inches, Pt

def open_file(filename):
  if sys.platform == "win32":
    os.startfile(filename)
  else:
    opener = "open" if sys.platform == "darwin" else "xdg-open"
    subprocess.call([opener, filename])

# From: https://stackoverflow.com/questions/287871/how-to-print-colored-text-to$
class bcolors:
  HEADER = '\033[95m'
  OKBLUE = '\033[94m'
  OKCYAN = '\033[96m'
  OKGREEN = '\033[92m'
  WARNING = '\033[93m'
  FAIL = '\033[91m'
  ENDC = '\033[0m'
  BOLD = '\033[1m'
  UNDERLINE = '\033[4m'

screenshotsDirectory = '/Users/jesusdelvalle/Documents/projects/make_slides/screenshots/'

nameOfExcelColumnWithURLs = 'URL'
nameOfExcelColumnWithCompanyName = 'Name'
nameOfExcelColumnWithYearFounded = 'Founded'
nameOfExcelColumnWithHQCity = 'HQ'
nameOfExcelColumnWithDescription = 'Description'

# This title will appear on the top of every slide.
title = "Top30 BRITISH / WHERE to Make an Internship 2021"

powerpoint = Presentation("My_Template_16x9_2.pptx")
finalPowerPoint = "my_slides_2.pptx"

# The following measures are in inches.
# titleHeight is the height reserved for the title.
# The title font size could be smaller.
# The total size is 16 x 9.
totalHeight = 9
titleHeight = 1
referenceHeight = 3
pictureHeight = Inches (referenceHeight)
pictureWidth = Inches (referenceHeight*9/5)
# Total height is 9 inches. Reserving titleHeight for the title,
# nRowsPerSlide = (9 - titleHeight) // referenceHeight
# In Python, the ( // ) operator returns the integer quotient.
nRowsPerSlide = (totalHeight - titleHeight) // referenceHeight

try:
  # Expected: Excel file with ending .xslx
  inputFilename = sys.argv[1]

  excel = pd.read_excel(inputFilename, header=0)

  # i is the counter for the number of items which will go into the slides.
  i = 0
  # iRow is the counter for the number of rows which will go into each slide.
  iRow = 0
  layout = powerpoint.slide_layouts[0]
  slide = powerpoint.slides.add_slide(layout)
  shape = slide.shapes
  for index, row in excel.iterrows():

    url = row [nameOfExcelColumnWithURLs]

    url = url.replace("https://","")
    url = url.replace("/","")
    print (url)
    pictureFile = screenshotsDirectory + url + ".png"

    if iRow == 0:
      slide = powerpoint.slides.add_slide(layout)
      textBox = slide.shapes.add_textbox (Inches (0.05), Inches (0.0), Inches(14), Inches(1))
      textFrame = textBox.text_frame
      textFrame.word_wrap = True
      # A text_frame always contains one paragraph.
      p = textFrame.paragraphs[0]
      p.font.size = Pt(40)
      p.font.color.rgb = RGBColor(245, 223, 77)
      p.text = title

    # 16:9 -> 3 columns each 5.33, 4 rows each 2.25
    # new height = (5/9) * new width

    # 13.3 x 7.5 inches -> 4.43 x 1.875

    left = Inches (0)
    top = Inches (titleHeight + (iRow * referenceHeight) + (iRow * titleHeight / 3))

    picture=slide.shapes.add_picture (pictureFile, left, top, width = pictureWidth, height = pictureHeight)

    # Add text.
    # From https://python-pptx.readthedocs.io/en/latest/user/text.html

    j = 0

    txBox = slide.shapes.add_textbox (left + pictureWidth, top, Inches(8), Inches(3))
    tf = txBox.text_frame
    tf.word_wrap = True

    p = tf.add_paragraph()
    p.font.size = Pt(40)
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.text = row [nameOfExcelColumnWithCompanyName]
    p = tf.add_paragraph()

    # Clean missing values.
    rowFounded = ""
    if pd.isnull (row [nameOfExcelColumnWithYearFounded]):
      rowFounded = ""
    else:
      rowFounded = "Founded: " + str (int (row [nameOfExcelColumnWithYearFounded])) + ". "
    p.font.size = Pt(20)
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.text = row [nameOfExcelColumnWithHQCity] + '. ' + rowFounded + row [nameOfExcelColumnWithDescription]

    p = tf.add_paragraph()
    p.font.size = Pt(12)
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.text = row [nameOfExcelColumnWithURLs]

    j = j + 1
    # print ("Shape written: " + str (j))
    iRow = iRow + 1
    # print ("iRow: " + str(iRow))
    # iRow is 0 indexed.
    if iRow == nRowsPerSlide:
      iRow = 0
    i = i + 1
  print ("Number of companies: ", i)
  powerpoint.save (finalPowerPoint)

  # Open the presentation file.
  # From https://stackoverflow.com/questions/17317219/is-there-an-platform-independent-equivalent-of-os-startfile/17317468#17317468

  open_file (finalPowerPoint)
except IOError:
  print(f"{bcolors.FAIL}IO Error (Wrong filename?): {bcolors.ENDC}" + inputFilename)
except IndexError as error:
  print(f"{bcolors.FAIL}Index Error (Forgot filename?):{bcolors.ENDC}", error)
except Exception as error:
  print(f"{bcolors.FAIL}Error:{bcolors.ENDC}", error)
