Working with map files 

Under the folder input_maps, you will see a file with a .ini extention which can be opened with a text editor.
This file can be edited to map the approprite coloumn name on the excel file to the right text box on the pdf file 


HOW TO FIND NAMES OF THE TEXT BOXES
1. First, move the new pdf form to this directory which contains the script and rename it to something that can be addressed easily 

2. Open the script "textboxmapper.py" on the terminal

3. The script will prompt you to enter the path to a pdf file. Enter the name of the file (including .pdf) and press enter 

4. When the script exits, you will find two new files named <original_file_name>-mapped.pdf and <original_file_name>.txt 

5. Open the newly created pdf file, you will find that the textboxes are filled with their corresponding names

6. To know the names of checkboxes in the pdf form, consult the text file with the same name, it contains all the editable elements
   on the pdf in order. Simply look for the textboxes that sorround the checkbox and look for an element in the text file
   that's not represented on the mapped pdf 

WORKING WITH MAP FILES 

1. The map file with the extention .ini under the folder input_maps is the main part of the script.

2. Open the map file on a text editor, you will see a bunch of settings with 3 sub-divisions 


SETTINGS 

1. 'Settings' is the first subdivision in the file. It contains vital information to run the script. 

2. The value 'source_file' points to the excel file that needs to be read from 

3. The value 'pdf_form' points to the main pdf form that needs to be filled 

4. The value 'pdf_form2' points to the supplementry pdf form to be filled 

5. The value 'base_map' points to the worksheet on the excel file that will be used for the main form 

6. the value 'identifier' points to the column that can be used to identify unique rows 

CHANGING THE MAPPING 

1. Map the excel file to the pdf form by equating the coloumn name with its appropriate text box name. 

2. The column names are not case sensetive but the text box names are case sensetive 



  