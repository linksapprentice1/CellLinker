CellLinker
==========

Dually link excel cells!

Edit cells_to_link.txt to specify what cells you wish to be linked. 
Cells are represented by an address and sheet name, delimited by space. 
Cells themselves are delimited by space.
Each row represents a set of cells that will be linked to each other
(i.e. if you change the value of any one, it will change the rest to that value).

Run the Python code. A text file will be generated containing VBA code.

Open an Excel file. Developer -> Visual Basic -> ThisWorkbook (in tree on left, under Microsoft Excel Object, under VBAProject). Copy and paste the VBA code into the editor. 
Close out of the file and reopen it. It will now function properly.

