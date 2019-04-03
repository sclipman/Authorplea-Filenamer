###############################################################################
# Program: authorplea
# Type: Python 3 Script
# Author: Steven J. Clipman
# Description: This program prefixes all .docx files in the given directory
# with the last name of the user who authored it.
# Usage: Place all .docx files to rename in a directory and execute script.
# Example: python3 ~/scripts/authorplea.py
###############################################################################


import zipfile, lxml.etree
import os

def getFiles(path):
    files = []
    for (dirpath, dirnames, filenames) in os.walk(path):
        files.extend(filenames)
        break
    return(files)

def getAuthor(filename):
    zf = zipfile.ZipFile(filename)
    doc = lxml.etree.fromstring(zf.read('docProps/core.xml'))
    ns={'dc': 'http://purl.org/dc/elements/1.1/'}
    creator = doc.xpath('//dc:creator', namespaces=ns)[0].text
    return(creator)

def main():
    path = input("Enter folder path: ")
    os.chdir(path)
    files = getFiles(path)
    for file in files:
        if file[-5:] == '.docx':
            try:
                author = getAuthor(file).split()[1]
                rename = author+'_'+file
                os.rename(file,rename)
            except:
                print("Error:", file, "did not have proper author metadata.")
    print("\nFiles in have been renamed!\n")
main()
