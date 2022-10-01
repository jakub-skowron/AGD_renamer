# AGD_renamer

### Description

AGD_renamer was created to automate the files renaming process in according to the parts list. Parts file's names organisation is necessery to making documentation in our production plant. Program gets data from the parts list, prepares a correct words array and changes files names (in .png, .jpg, .dxf, .stp formats).

### Requirements

For python 3.X users just execute:

`pip install -r requirements.txt`

### How to run this program?

The program can be run by: `python AGD_renamer.py`

### What I learned
In this project I learned how to get information from excel using openpyxl library and how to create a graphical user interface with tkinter package.

### Example of use
1. Before start, make sure that:
- The excel file is based on the added example - in this special standard - and is in .xlsx format.

![image](https://user-images.githubusercontent.com/104097333/193411109-c14d40f4-73af-46cd-acb9-895dbf1f3d22.png)

- Each file name starts from the position in the parts list (example below). If not program doesn't change file's name.

![image](https://user-images.githubusercontent.com/104097333/193411974-ac746395-c987-474a-9885-a2bb65be7f15.png)

2. User should select the excel file with parts list and the path to the files to be renamed. 

![image](https://user-images.githubusercontent.com/104097333/193151584-6b78aa25-89bf-45fd-a8b2-cc8c33989f3c.png)

![image](https://user-images.githubusercontent.com/104097333/193411198-c8c0d892-a15b-4d85-a48a-0ebb97313dc0.png)

![image](https://user-images.githubusercontent.com/104097333/193411232-35cbbda8-4b8c-416f-975e-d2335067081c.png)

3. When you choose paths, click Run button.

![image](https://user-images.githubusercontent.com/104097333/193417504-4a45c49f-bf48-42ca-aaec-42c85000cc03.png)

4. Program should show you how many files names were changed.

![image](https://user-images.githubusercontent.com/104097333/193411626-dca63e93-3f8b-4085-976e-ea8fc98ef8ca.png)

![image](https://user-images.githubusercontent.com/104097333/193417650-e0a5ea40-d497-4ee6-8710-026ae00c49bf.png)
