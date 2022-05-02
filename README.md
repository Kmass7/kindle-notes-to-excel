# kindle-notes-to-excel
Takes "My Clippings" text file from an Amazon Kindle device and formats information in excel files based on book

This allows you to easily highlight something on your Kindle, maybe write a short note with its shoddy keyboard system, and then write something more fleshed out in Excel if you wish. It also just makes it much nicer to look and read.

The clippings file is really just a "note history" file, so even if you delete a highlight/note in your kindle... it will still show up in the file. Take for example you highlight a small portion of the text, but the later you want to expand that selection. The kindle will save that bigger selection as a separate clipping rather than just overwriting the old one. You will still have to remove duplicates from time to time as the code can not read your mind and know which one you rather keep.

General things to do:
- Add your paths to the fileLocations.py file

- Make sure your kindle is plugged in

- Do not have any of the excel note files open, or else you will be prompted to close

- If you want, add the format_cells.bas macro to your excel PERSONAL.XLSB modules. Use the shortcut Ctrl+Shift+F to format the cells into a readable manner.

