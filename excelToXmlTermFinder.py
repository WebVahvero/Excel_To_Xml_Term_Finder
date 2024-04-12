import xml.etree.ElementTree as ET
import pandas as pd
import json

# FILENAMES AND PATHS
xmlFileNameWithPath = ""
excelFileNameWithPath = ""

# TAG IN XML WHICH IS HOLDING THE VALUE TO BE REPLACED WITH TERM FROM THE EXCEL
targetTag = ""

# VARIABLES TO SAVE SOME INFORMATION ON PASSED TARGET TAGS
wordsToBeTranslated = []
targetRowCountTotal = 0
targetRowCountPassedFilter = 0

# XML FILE READING AND PARSING
tree = ET.parse(xmlFileNameWithPath)
root = tree.getroot()

# LOOPING XML FILE WITH "target" TAG AND EXLUDING WORDS WITH SPECIAL CHARACTERS
for child in root.iter(targetTag):
    targetRowCountTotal = targetRowCountTotal + 1
    if "{" in child.text or "-" in child.text or "&" in child.text or "Col" in child.text or "<" in child.text or ":" in child.text or "#" in child.text or "_" in child.text:
        continue
    else:
        wordsToBeTranslated.append(child.text)
        targetRowCountPassedFilter = targetRowCountPassedFilter + 1
          
# EXCEL FILE READING AND PARSING
TerminologyExcel = pd.read_excel(excelFileNameWithPath, skiprows=3)

# SAVING COLUMNS TO VARIABLES BY THE COLUMN NAME
fiColumn = TerminologyExcel["FI"].values
enColumn = TerminologyExcel["EN"].values
fiList = list(fiColumn)
enList = list(enColumn)

# VARIABLES FOR HOLDING INFO ON MATCHING WORDS FOUND IN EXCEL AND XML TAGS
foundWords = 0
unFoundWords = 0
amountOfWordsWithTranslation = 0

# Collecting unfound words to list
# Collecting found words to dict keys and terms to values
wordsWithoutTranslation = []
wordsWithTranslation = {}

# COMPARING XML FILE WORDS WITH "EN" COLUMN IN EXCEL FILE AND THEN FINDING IF THERE IS TRANSLATION IN "FI" COLUMN
# IF TRANSLATION FOUND ADDING IT IN DICTIONARY WITH ENGLISH WORD IF NOT ADDING ENGLISH WORD TO LIST WITHOUT TRNSLATION
for word in wordsToBeTranslated:
    if word in enColumn:
        foundWords = foundWords + 1
        if pd.notnull(fiList[enList.index(word)]) != True:
            wordsWithoutTranslation.append(word)
        else:
            #print(f"{word} = {fiList[enList.index(word)]}")
            wordsWithTranslation[word] = fiList[enList.index(word)]
            amountOfWordsWithTranslation = amountOfWordsWithTranslation + 1
    else:
        unFoundWords = unFoundWords + 1
        
print(f"\nTotal amount of target tags {targetRowCountTotal}\nTarget tags with actual words {targetRowCountPassedFilter}\n")
print(f"Cell found in excel with word to be translated to xml {foundWords}\nNo cell from excel with the word to be translated found {unFoundWords}\n")
print(f"Words with translation {amountOfWordsWithTranslation}\n")

#print(json.dumps(wordsWithTranslation, indent=4))
#print(*wordsWithoutTranslation, sep = "\n")

tagsUpdated = 0

# FINDING AND REPLACING TRANSLATEABLE WORDS WITH WORDS FROM THE "FI" COLUMN OF THE EXCEL BETWEEN "target" TAGS
for targetWord in root.iter(targetTag):
    if "{" in targetWord.text or "-" in targetWord.text or "&" in targetWord.text or "Col" in targetWord.text or "<" in targetWord.text or ":" in targetWord.text or "#" in targetWord.text or "_" in targetWord.text:
        continue
    else:
        if targetWord.text in wordsWithTranslation.keys():
            for key, value in wordsWithTranslation.items():
                if key == targetWord.text:
                    #print(f"Target \"{targetWord.text}\" found with translation \"{value}\"")
                    targetWord.text = value
                    tagsUpdated = tagsUpdated + 1
                else:
                    continue
                
tree.write('output.xml', encoding='UTF-8')
print(f"{tagsUpdated} Target tags updated with translation\n")