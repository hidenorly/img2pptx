#   Copyright 2023 hidenorly
#
#   Licensed under the Apache License, Version 2.0 (the "License");
#   you may not use this file except in compliance with the License.
#   You may obtain a copy of the License at
#
#       http://www.apache.org/licenses/LICENSE-2.0
#
#   Unless required by applicable law or agreed to in writing, software
#   distributed under the License is distributed on an "AS IS" BASIS,
#   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#   See the License for the specific language governing permissions and
#   limitations under the License.

import argparse
import os
import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt

class PowerPointUtil:
    SLIDE_WIDTH_INCH = 16
    SLIDE_HEIGHT_INCH = 9

    def __init__(self, path):
        self.prs = Presentation()
        self.prs.slide_width  = Inches(self.SLIDE_WIDTH_INCH)
        self.prs.slide_height = Inches(self.SLIDE_HEIGHT_INCH)
        self.path = path

    def save(self):
        self.prs.save(self.path)

    def addSlide(self, layout=None):
        if layout == None:
            layout = self.prs.slide_layouts[6]
        self.currentSlide = self.prs.slides.add_slide(layout)

    def addPicture(self, imagePath, x=0, y=0, width=None, height=None, isFitToSlide=True):
        pic = self.currentSlide.shapes.add_picture(imagePath, x, y)
        if width and height:
            pic.width = width
            pic.height = height
        else:
            if isFitToSlide:
                width, height = pic.image.size
                if width > height:
                    pic.width = Inches(self.SLIDE_WIDTH_INCH)
                    pic.height = Inches(self.SLIDE_WIDTH_INCH * height / width)
                else:
                    pic.height = Inches(self.SLIDE_HEIGHT_INCH)
                    pic.width = Inches(self.SLIDE_HEIGHT_INCH * width / height)

    def addText(self, text, x=Inches(0), y=Inches(0), width=None, height=None, fontFace='Calibri', fontSize=Pt(18), isAdjustSize=True):
        if width==None:
            width=Inches(self.SLIDE_WIDTH_INCH)
        if height==None:
            height=Inches(self.SLIDE_HEIGHT_INCH)

        textbox = self.currentSlide.shapes.add_textbox(x, y, width, height)
        text_frame = textbox.text_frame
        text_frame.text = text
        font = text_frame.paragraphs[0].font
        font.name = fontFace
        font.size = fontSize
        theHeight = textbox.height
        
        if isAdjustSize:
            text_frame.auto_size = True
            textbox.top = y


if __name__=="__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--input", default=".", help="Input folder path")
    parser.add_argument("-o", "--output", default="output.pptx", help="Output PowerPoint file path")
    parser.add_argument("-f", "--addFilename", default=False, action='store_true', help="Add filename to the slide")
    args = parser.parse_args()

    prs = PowerPointUtil( args.output )

    imgPaths = []
    for dirpath, dirnames, filenames in os.walk(args.input):
        for filename in filenames:
            if filename.endswith(('.png', '.jpg', '.jpeg', '.JPG')):
                imgPaths.append( os.path.join(dirpath, filename) )

    imgPaths.sort()

    for filename in imgPaths:
        prs.addSlide()
        prs.addPicture(os.path.join(dirpath, filename), 0, 0)
        if args.addFilename:
            prs.addText(os.path.basename(filename), Inches(0), Inches(PowerPointUtil.SLIDE_HEIGHT_INCH-0.4), Inches(PowerPointUtil.SLIDE_WIDTH_INCH), Inches(0.4))

    prs.save()
