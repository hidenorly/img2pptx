# img2pptx


# requirements

```
pip3 install python-pptx
pip3 install pip pillow
brew install libffi libheif
pip3 install git+https://github.com/david-poirier-csn/pyheif.git
```

# how to use

```
python3 img2pptx.py --help
usage: img2pptx.py [-h] [-i INPUT] [-o OUTPUT] [-f]

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT, --input INPUT
                        Input folder path
  -o OUTPUT, --output OUTPUT
                        Output PowerPoint file path
  -f, --addFilename     Add filename to the slide
```

```
$ python3 img2pptx.py -i ~/imageFolder -o ~/output.pptx -f
```

Then all of image files in ~/imageFolder are read and output to the ~/output.pptx and each image has filename in the bottom of each slide.