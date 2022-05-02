import fileLocations
import re
import json
import pandas as pd
from collections import defaultdict
from psutil import process_iter
from openpyxl import load_workbook
from pathlib import Path

storageFile = Path(fileLocations.srcPath) / "clips.json"

termLine = '=========='

def getClipping(filename):
    #read My Clippings file and remove the 10 separating equals signs
    with open(filename, 'r', encoding='utf-8-sig') as f:
        sourceFile = f.read()
    return sourceFile.split(termLine)


def getRows(clipping):
    #segment each clipping into basic information
    row = {}
    lines = [l for l in clipping.split(u'\n') if l]

    #'lines' should return a 3 element list (empty lines are actually written as '\n')
    if len(lines) != 3:
        return

    firstLineSplit = lines[0].split(' (')
    row['Title'] = str(firstLineSplit[0])
    row['Author'] = str(firstLineSplit[1]).replace(')', '') #Keep the Author Key in case you want to create folders for each Author later on down the line

    #Selected Kindle Passages are listed as "Highlights"
    if "Highlight" in lines[1]:
        #Match a range of locations
        row['Type'] = 'Highlight'
        match = re.search(r'(\d+)-(\d+)', lines[1])
        startLocation = match.group(1)
        endLocation = match.group(2)

    #Written comments in the device are listed as "Notes"
    elif "Note" in lines[1]:
        #Match a single number location, with size Zero
        row['Type'] = 'Note'
        match = re.search(r'Location (\d+)', lines[1])
        startLocation = match.group(1)
        endLocation = startLocation

    #Bookmarks are listed in the "My Clippings.txt" but they are not needed
    elif "Bookmark" in lines[1]:
        return

    row['Start Location'] = int(startLocation)
    row['End Location'] = int(endLocation)
    row['content'] = lines[2]

    return row


def loadClips():
    #load old clips
    try:
        with open(storageFile, 'r') as f:
            return json.load(f)

    #If it doesn't exist yet
    except (IOError, ValueError):
        return {}


def saveClips(myClippings):
    #save new clips
    with open(storageFile, 'w') as f:
        json.dump(myClippings, f)


def hasHandle(fpath):
    #Ripped from StackOverflow - could be changed to be faster, but I'm not too sure how to do it yet lol
    #See if process is running
    for proc in process_iter():
        try:
            for item in proc.open_files():
                if fpath == Path(item.path):
                    return True
        except Exception:
            pass

    return False


def myClippingsExcel(newLineDict):
    #Create new workbooks if they don't exist
    for title in newLineDict:

        #Remove forbidden file name characters, set path
        writeFile = str("%s.xlsx" % title)
        writeFile = re.sub(r'[\\/*?:"<>|]',"", writeFile)
        writeDir = Path(fileLocations.outputPath) / writeFile

        if Path(writeDir).exists() == False:
            fields = {'Location':[],"Passage":[],'Notes':[],'Other Comments':[]}

            with open(writeDir, 'wb') as f:
                df = pd.DataFrame(fields)
                writer = pd.ExcelWriter(f, engine='xlsxwriter')
                df.to_excel(writer, sheet_name='Sheet1', index=False)
                workbook  = writer.book
                macroFile = re.sub(r'[\\/*?:"<>|]',"", title)
                workbook.filename = f'{macroFile}.xlsm'
                workbook.add_vba_project(Path(fileLocations.vbaProjectPath))
                writer.save()
        
        Path.unlink(writeDir)
        print(f"{writeFile} was created.")

    #Update existing  workbooks
    for title in newLineDict:

        #Remove forbidden file name characters, set path
        writeFile = str("%s.xlsm" % title)
        writeFile = re.sub(r'[\\/*?:"<>|]',"", writeFile)
        writeDir = Path(fileLocations.outputPath) / writeFile

        openTest = hasHandle(writeDir)

        while openTest == True:
            x = input(f"The Excel File for {writeFile} is open\nPlease close the file and press enter to continue!")
            if hasHandle(writeDir) == False:
                break
            

        wb = load_workbook(writeDir, read_only=False, keep_vba=True)
        ws = wb.worksheets[0]

        for content in newLineDict[title]:
            # Append Row Values
            ws.append(newLineDict[title][content])
            wb.save(writeDir)

        print(f"{writeFile} was updated")


def main():
    #Setup dicts for formating Kindle Clipping segments into usful forms
    myClippings = defaultdict(dict)
    oldClips = defaultdict(dict)
    newLineDict = defaultdict(dict)
    oldClips.update(loadClips())
    if len(oldClips) > 0:
        print("Load successful")

    else:
        print("No stored clippings loaded")

    #For separating "Highlights" from "Notes"
    clippingList = []
    noteList = []

    #Only works when Kindle is plugged in
    clipFile = getClipping(Path(fileLocations.myClippingsPath))

    #Retrieve relevent clipping information (Author/Book/Location/Type/content)
    #Sort Types into separate lists
    for clipping in clipFile:
        row = getRows(clipping)
        if row and row.get('Type') == 'Highlight':
            clippingList.append(row.copy())
        if row and row.get('Type') == 'Note':
            noteList.append(row.copy())

    #Create a new dictionary Key, "Note", with the value of a matching clipping note's "content" value
    for note in noteList:
        for highlight in clippingList:

            #The ideal clipping has the note's location match the highlights END location
            if highlight.get('End Location') == note.get('End Location'):
                highlight['Note'] = note.get('content')
            
            #If the highlight was changed to go further than where the note is located.
            elif highlight.get('End Location') > note.get('End Location') and highlight.get('Start Location') < note.get('End Location'):
                highlight['Note'] = note.get('content')

            #What about if the highlight location shrinks, but the note stays where it was? IDK, just try to not do that :P.
            
            #Format
            myClippings[highlight['Title']][highlight['content']] = [int(highlight.get('End Location')), str(highlight.get('content')),str(highlight.get('Note'))]
        
    #Get rid of empty values
    myClippings = {k: v for k, v in myClippings.items() if v}

    #Prevent overwriting; check saved clips and keep only the new clips
    for title in myClippings:
        for content in myClippings[title]:
            if content not in oldClips[title]:
                newLineDict[title][content] = myClippings[title][content]
    

    #Prevent failure at first use. (DO NOT DELETE clips.json!!!) 
    if len(oldClips) == 0:
        saveClips(myClippings)
        
    if len(newLineDict) > 0:
        print("Adding new clips...")

    myClippingsExcel(newLineDict)
    saveClips(myClippings)


if __name__ == "__main__":
    main()
