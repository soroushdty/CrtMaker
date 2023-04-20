Welcome to Certificate Maker User Manual
This script automatically issues certificates from a PSD template file with names extracted from a MS Excel file.
Certificate Maker is written in Python 3.10.7. and runs on Python 3.7 and newer.
Note 1: Certificate Maker only runs on Windows OS.
Note 2: Certificate Maker does not change fonts, formatting, etc. Make sure to install all necessary fonts and format the layer with your desired font, size, color, etc.
<img src="../Screenshots/master/1.png">
How to use Certificate Maker:
1. Install Python on your machine, if not already installed. You can find the latest Python version here: www.python.org/downloads. You must also have Photoshop installed.
2. Download the program (CrtMaker.py). I strongly recommend putting the program file in the same directory as both the PSD template and the excel file.
3. Run the program with an active internet connection.
If you get an error saying “Python is not recognized as an internal or external command”, do the following:
Find the folder where Python is installed: Go to Start à Python à Open file location à Properties and copy the “Starts in:" field
Right-click “This PC”, then go to Properties à Advanced system settings à Environment variables.
In the window that appears, if a “path” variable exists, select it, ​and click Edit; otherwise, click New.
In the next dialogue box, click on New and paste the previously copied path of the folder; then, click OK.
4. The first run might take a while as the program will download and install the needed packages from the internet. The program will automatically open Photoshop. Remember to minimize the Photoshop window; Do not close it and don’t touch it as it might break the program!
5. If everything goes fine, you should see the following message:

6. If you have put the PSD file in the same directory as the program, press 1 then Enter, otherwise enter the full address in the following format: X:\...\For example:  C:\Users\John\Desktop\
7. A list of PSD files in the selected directory is shown. Enter the number you want to select that file

8. A list of layers in the PSD file is shown. Enter the number you want to select that layer:

9. Select where your Excel file is located. Press 1 to use the same category as the PSD file or enter the full address in the following format: X:\...\

10. A list of Excel files in the selected path is shown. Choose your file using its number

11. If you excel has more than one sheet, the program will ask you to choose a sheet

12. Select the column that contains the names using a capital letter from A to Z

13. Select the first and the last row that contain the names in that column:

14. The program will show a summary of selected names and ask you where you want your certificates saved. Press 1 to save them in the same path as the PSD file or enter the full address in the following format: X:\...\Note that the program will automatically create a folder called “Certificates” in that location.

15. The program start issuing certificates and will exit automatically when all certificates are issued. The files will be saved as PNG images in your selected category in a folder called Certificates (or Certificates 2/Certificates 3 if a Certificates folder already exists).

17. Exit Photoshop when you are finished. Do not save anything.

*********************************************************************************************
Feel free to contact me for any inquiries: soroushdianaty@gmail.com
Certificate Maker is distributed under GNU General Public License v3.0 license.
