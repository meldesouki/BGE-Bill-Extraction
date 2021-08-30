# BGE Bill Extraction

## Overview 
This program will allow you to quickly extract certain fields from a machine-generated BGE bill and export the data into an Excel file of your choosing. It is designed to work with the Flight Plan Excel template and will automatically create new columns for each bill scanned. The files in *BGE Bill Extraction.zip* (which contains the GUI) is only compatible with Windows. However, the scripts are cross-platform compatible. 

## How to use it
1. Simply download the file *BGE Bill Extraction.zip* and unzip it. 
2. Click on the shortcut called *BGE Bill Extraction* which will launch the program. A command prompt may open as well however this can be ignored. Do *not* close it or else the program will stop running.
3. Choose a bill to extract from and follow the prompts on screen.
4. Choose the Excel file you would like the data to be saved to and click "OK". The Excel file *cannot* be open in the background as the program is running.
5. Once the program is done, a popup window will appear saying "Extraction completed successfully".  

## Restrictions
Since this program doesn't scan bills with OCR or use machine learning there are a few restrictions on what this program can do.

- It can only read from bills that are machine-generated PDF files. If they are saved in any other format or are not machine-readable (meaning you can't copy and paste from the file), this won't work
- The PDF must be a BGE bill and it must not have any abnormal formatting. For example, some bills don't have green headers at the top of each page. Those bills are not compatible with this program.

    NOTE: Bills that have multiple usage totals for the same commodity are compatible. Other anomalies are not. 

## Acknowledgements
This program uses the [PDFplumber](https://github.com/jsvine/pdfplumber) library by jsvine and the [PySimpleGUI](https://github.com/PySimpleGUI/PySimpleGUI) library