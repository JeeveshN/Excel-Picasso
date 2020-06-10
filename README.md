# Excel-Picasso
A fun side project    
Takes any picture and paints it in Excel.

## Installation
```
pip install -r requirements.txt
```

## Usage
Clone the repository and run painter.py -h
```
usage: painter.py [-h] --image IMAGE [--orignal] [--tiny]

optional arguments:
  -h, --help     show this help message and exit
  --image IMAGE  Path to image file
  --orignal      Keep the orignal size of the image.Without this image would
                 be scaled down
  --tiny         Scale down image to 12 percent.Usefull while dealing with
                 very high resolution images
```

## Voila
![Alt text](/extras/example.jpg?raw=true "Orignal Image")
![Alt text](/extras/example.png?raw=true "Painted Spread Sheet")

*Note - The picture is visible best at the lowest pissble zoom level. The example shows red lines across, which is due tue     github compressing the image, the spreadsheets have been included in the extras folder feel free to download and play around works best if you open these is **MS Excel***
