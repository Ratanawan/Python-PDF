# Python-PDF

This project is to read questions and highlighted answer from pdf file into word document

For example: You have a PDF file from your lecturer which contains QCM for exam. You want to extract only the question and answer into a document so that you can memorize them faster. Hence, you can use this program to extract those information like you desire.

Modules to Install:
1. pip install os
2. pip install regex
3. pip install docx
4. pip install pyperclip

Steps to follow: 
1. Change your directory path to your desire path that you want to save in the Line 6
2. Change directory path in the Line 58 to your PDF location
3. Change the name of Word document that you want to save in the Line 74
4. Save the Python File
5. Open your PDF, Select all contents in the PDF and Copy (Ctrl + a & Ctrl + c)
6. Run the python program
7. You can find your saved word document in the path that you saved from the first step

Example PDF file:
![PDF](https://user-images.githubusercontent.com/34526877/117576987-17b1fb80-b112-11eb-8757-62e68c5a578f.PNG)

Example Output in docx: 
![Doc](https://user-images.githubusercontent.com/34526877/117577026-4d56e480-b112-11eb-8ae8-cc425d06117b.PNG)

For reading the highlighted text part, I got the code from https://stackoverflow.com/questions/9099497/how-to-extract-highlighted-parts-from-pdf-files
