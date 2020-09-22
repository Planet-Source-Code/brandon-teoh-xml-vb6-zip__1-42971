This sample provides demonstration of XML using MSXML4 and VB6.

It is capable of the following:

1. Open XML document.
2. Create XML document.
3. Edit XML document.

The concept for 'Create Document' and 'Edit document' is slightly different, which gave rise to redundant codes. However, both
ways have its pros and cons. One will discover this when trying out the functions.

What you need ?
----------------
1. MSXML4 (http://download.microsoft.com/download/xml/SP/40SP1/WIN98MeXP/EN-US/msxml.msi)
2. VB6 with at least SP2 (for TreeView control) 

Limitations
------------
1. While deleting an node from the treeview, bubblesort has to be performed. Thus if the branches are deep, it will take a while.
2. Functions cater for DTD standard do not work for XML-schema standard.

Procedures
---------------
1. Install MSXML4 first
2. Goto 'XML_Dtd' folder, copy 'DB.dtd' and 'notes2.dtd' into 'c:'. Failing this, certain function might not work.
3. Run the program(test.exe), refer to manual.doc for basic instruction.

Other Contributors
-------------------
1.Lamont Adams - of http://builder.com.com
2.'Coolwick' of planet-source-code
3. 'BelgiumBoy' - of http://www.bartnet.freeservers.com/

Reference
---------
1. http://www.w3schools.com/dtd/default.asp

Conclusion
-----------
XML document is still best created manually using excel or other editor.

brandonteohno1@yahoo.com