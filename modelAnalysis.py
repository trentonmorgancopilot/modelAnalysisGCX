"""
Walk through each element (class, table, view, etc.) in the rsmGCX model and use Agent Ransack to inspect references.
"""
import os, sys
import subprocess
import json
from xml.etree import ElementTree as ET
from pprint import pp as pprint

MODEL_DIR = "K:\\AosService\\PackagesLocalDirectory\\rsmGCX\\rsmGCX"
RANSACK_PATH = '"C:\\Program Files\\Mythicsoft\\Agent Ransack\\AgentRansack.exe"'
RESULTS_PATH = "C:\\Users\\Admin56a04fd4cc\\Desktop\\ModelAnalysisGCX\\RansackResults.txt"
OUTPUT_PATH = "C:\\Users\\Admin56a04fd4cc\\Desktop\\ModelAnalysisGCX\\ModelAnalysisOutput.txt"

"""
Primary logic function.
"""
def main():

    # Init dict to store all elements and their references
    # Dictionary contains pairs {<elementName> : <references>}
    references = {}

    # Init list to track elements that require manual inspection
    inspectList = []

    # Init and populate dictionary to store XML file paths
    # Dictionary contains pairs {<filePath> : <elementName>}
    elements = getXMLFilePaths(MODEL_DIR)
    
    # For each XML file, find the value of the <Name> tag, then find and inspect all references using Agent Ransack
    for filePath in elements:
        elementName = parseXML(filePath)
        
        if elementName is not None:
            try:
                print("ELEMENT NAME: " + elementName + "\n------------")
            
                # Since some code can exist without being referenced in other files, DO NOT process elements that have
                # names ending in "_Extension", ".rsmGCX", or ".GCX"
                if (elementName.endswith("_Extension")
                 or elementName.endswith(".rsmGCX")
                 or elementName.endswith(".GCX")
                    ):
                    print("Skipping element...")
                    print("+++++++++++\n")
                
                else:
                    # Agent Ransack is called directly within <findReferences()>
                    elements[filePath] = elementName
                    findReferences(elementName)
                    
                    # After search results are generated, retrive them from the output file
                    with open(RESULTS_PATH, 'r') as file:
                    
                        # Store the results as a list so we can analyze them
                        references[elementName] = file.readlines()
                    
                        # Any elements that have only one reference (itself) should be tracked
                        numberFileReferences = calcNumReferences(references[elementName])
                        
                        if numberFileReferences < 2:
                            inspectList.append(elementName)
                            print("Element added to inspection.")
                            
                        print(f"Number File References: {numberFileReferences}\n")
                        
                        # Reset file cursor so we can use <read()> below
                        file.seek(0)
                        
                        # Print references for readability
                        print(file.read().rstrip(), '\n')
                        print("+++++++++++\n")
            
            # If we catch an error, mark that element for review
            except Exception as e:
                print(f'Error finding references: {e}')
                inspectList.append(elementName)
                print("Element added to inspection.\n")
                print("+++++++++++\n")
    
    # Output the list of elements that need manual review
    print("INSPECT LIST:\n------------")
    pprint(inspectList)
        
    # After all elements have been processed, write the references dictionary to a file so we can reuse later
    with open(OUTPUT_PATH, 'w') as file:
        json.dump(references, file, indent = 4)
    
"""
Get the complete file path for each XML file within the given source <directory>.
Returns a dictionary with the paths as keys.
"""
def getXMLFilePaths(directory):

    xmlFilePaths = {}
    
    for root, directoryNames, fileNames in os.walk(directory):    
        if fileNames:
            for file in fileNames:
                if file.endswith(".xml"):
                    fullPath = (root + "\\" + file)
                    xmlFilePaths[fullPath] = {}
            
    return(xmlFilePaths)

"""
Parse the XML script in the given file and extract the value of the <Name> element.
Returns a string of the element text if it exists, otherwise returns None.
"""
def parseXML(filePath):

    tree = ET.parse(filePath)
    root = tree.getroot()

    value = tree.find('Name')
    if value is not None:
        return value.text
    else:
        print("No name found!")
        print("+++++++++++\n")
        return None

"""
Invoke Agent Ransack via CMD to find references to the given <elementName> within the GCX model.
Returns a list that contains each line of the search results as elements.
"""
def findReferences(elementName):

    # Build cmd string to search for element references
    # Documentation: https://help.mythicsoft.com/agentransack/en/commandline.htm
    # -c = Text to search for in specified files (i.e. the Containing Text field)
    # -d = Directory(s) to search (i.e. the Look In field), use -dw for current working directory
    # -s = Search subfolders
    # -o = Output filename (runs the search without showing the user interface, streaming results directly to the file)
    # -oa = Append to output file
    
    # Note that using "AgentRansack.exe" runs the Windows application instead of the console application equivalent,
    # "flpsearch.exe", which is not accessible in the free version of the program. If we do not output from this
    # command into a file using <-o> or <-oa>, the application will try to launch a GUI for every search. Consequently,
    # we must output directly to a file first, then read back in from that file to analyze the content.
    
    cmd = f'{RANSACK_PATH} -c "{elementName}" -d "{MODEL_DIR}" -s -o "{RESULTS_PATH}"'
    
    try:
        # Documentation: https://docs.python.org/3/library/subprocess.html#module-subprocess
        subprocess.run(cmd)
         
    except subprocess.CalledProcessError as e:
        print(f"Error executing Agent Ransack: {e}")

"""
Given some <references> as a list of lines generated by an Agent Ransack search, determine the number of distinct files
included in the output.
Returns an integer number of file references.
"""
def calcNumReferences(references):

    # Luckily, from the output it is very simple to tell how many separate files contain references. Below is an
    # example of the search results. Note that for each file path, there is a single list element containing
    # only the newline character '\n'. We need only to count the number of such elements present in the
    # results to determine the number of files containing references to the object.
    
    """
    ['K:\\AosService\\PackagesLocalDirectory\\rsmGCX\\rsmGCX\\AxClass\\rsmGCXActivityContractClass.xml '
     '1 KB XML File 4/8/2024 3:58:20 PM 4/8/2024 3:58:20 PM 4/8/2024 3:58:20 PM '
     '3\n',
     '3 \t<Name>rsmGCXActivityContractClass</Name>\n',
     '7 class rsmGCXActivityContractClass\n',
     '\n',
     'K:\\AosService\\PackagesLocalDirectory\\rsmGCX\\rsmGCX\\AxClass\\rsmGCXTimesheetWebservice.xml '
     '10 KB XML File 4/8/2024 3:58:23 PM 4/8/2024 3:58:23 PM 4/8/2024 3:58:23 PM '
     '3\n',
     "14     [ AifCollectionTypeAttribute('return', Types::Class, "
     'classStr(rsmGCXActivityContractClass))]\n',
     '17         rsmGCXActivityContractClass dataContract;\n',
     '23             dataContract = new rsmGCXActivityContractClass();\n',
     '\n']
    """
    
    numberFileReferences = references.count('\n')    
    return numberFileReferences

"""
Startup.
"""    
if __name__ == "__main__":
    main()
