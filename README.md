# Table of Contents
1. [File Organizer](#File-Organizer)
1. [PDF to CSV](#PDF-to-CSV)

# File Organizer
## Script
```File_Organizer.py```

## Packages Used
- hachoir 3.3.0
- Pillow 11.2.1
- pillow_heif 0.18.0

## Background
In 2025, I downloaded all the photos from my phone onto the computer. The software I used put all the photos and videos in one folder. To resolve this issue I created a python script to organize all ``` jpg, jpeg, png, HEIC, mov, mp4 and 3gp``` by the metadata ```Date taken```, for when the photo/video was taken. For example if a picture was taken in ```May 4 2019``` that file would be placed in the folder ```2019-05```. 

Because not all files have the ```Date taken``` metadata (i.e. screenshots), those files were placed in a folder called ```Unknown Date```. 

In the end this script was able to sort 5,633 files in under 30 seconds. Saving hours worth of work manually sorting files.


## Screen shots
The following image shows the unsorted files:
![Image of unsorted media files](Photos/File_Organizer_Unsorted.png?raw=True)

The following image shows the files sorted:
![Image of unsorted media files](Photos/File_Organizer_Sorted.png?raw=True)

# PDF to CSV

## Background
In 2020, I had a client who had a PDF file of people's contact info and wanted it to be in a CSV file.


## Packages
- XlsxWriter 3.2.2: 
    - pip install XlsxWriter
