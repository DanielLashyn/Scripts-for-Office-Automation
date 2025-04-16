# Table of Contents
1. [File Organizer](#File-Organizer)
1. [PDF to XLSX](#PDF-to-XLSX)

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

# PDF to XLSX

## Background
In 2020, I had a client, who was a teacher at the local Highschool, and had a PDF file of student's contact info and wanted it to be in a excel file. To save time, I created a python script to parse through the PDF and place the corresponding information in the excel file. The parsed information can be broken down into three sections:
- Student info (ID, Last/First/Middle Name, and Cell number)
- Guardian info 1 (Phone numbers, email)
- Guardian info 2 (Phone numbers, email)

## Screen shots
NOTE: The following images were done with 'mock data', due to the senstive nature of the original data.

The following image shows how the pdf file was setup
![Image of contact info in pdf](Photos/PDF_to_XLSX_Contact_PDF.png?raw=True)

The following image shows the files sorted:
![Image of contact info in excel](Photos/PDF_to_XLSX_Contact_xlsx.png?raw=True)
## Packages
- XlsxWriter 3.2.2: 
    - pip install XlsxWriter
