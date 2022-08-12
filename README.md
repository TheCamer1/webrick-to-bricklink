# webrick-to-bricklink
A program to convert webrick exports to bricklink xml files

# Instructions
To run this, replace 202003298_order_items.xlsx with your order excel file and edit WebrickConverter.py, replace the final line with print_bricklink_inventory('YOUR EXCEL FILE NAME HERE')

The program will search Rebrickable for items to ensure it gets the correct BrickLink part ID, and it will print out a list of items that it cannot find. It will then print out the BrickLink xml, which you can then copy paste into the BrickLink upload page. 
