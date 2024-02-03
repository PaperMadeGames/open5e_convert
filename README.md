# open5e_convert
A python script to extract data from open5e and format it in Docx files that are suitable for importing to Shard Tabletop

To use this script simply make sure all python packages are loaded and then cd into the folder and execute the script. (Better directions are coming)

Examples:
To convert all creature codex monsters to Docx for shard import:
python3 open5eConvert.py -f 'monsters' -s 'cc'

To convert all Monster Menaerie monsters to Docx for shard import:
python3 open5eConvert.py -f 'monsters' -s 'menagerie'

Once you have created a Docx file then you will need to go to shardtabletop.com and create a new book. When prompted you will want to "Import Book" and then select the docx file you generated with this script. Give your new book a name and click "create".

After that the book will open in editing mode. Click on the section you want to convert (e.g. 'Monsters'). Click the 'hamburger' button next to the section, then choose "Convert Children". Then choose the appropriate thing you are importing. Check that the content looks good and Click the "Keep and Update Book".

You are done. you have successfully imported stuff to shard.

A note on deleting... If you find that stuff didn't work you will want to delete all the spells created by the book by going to "my Creations" and searching for everything in the book you created. Then delete all the stuff you imported. AFTER you have deleted the things then you can delete the book and start over. If you delete the book first then you will need to manually delete all the things you imported.
