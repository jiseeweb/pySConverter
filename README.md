PySConverter

Updates MDB table of Biometric (ZKTeco and Granding) devices which can be use to upload in our in HR System. 


It simply updates the value of USERID of the CHECKINOUT table to its corresponding BADGENUMBER in USERINFO table and saves it to a specified folder.

User must specify the filepath of mdb file that they want to update, other fields can be ignored ( which has default value )
Save destination which defaults to desktop. It will create a folder and a subfolder where it will save thr updated mdb.:
Driver name defaults to Microsoft Access Driver ( *.mdb, *.accdb )
USERINFO Table name where we will get the badgenumber
Column and Table name to update separated by comma(,)
