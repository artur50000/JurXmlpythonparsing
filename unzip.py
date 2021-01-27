import zipfile
import os

working_directory = 'yourdir'
os.chdir(working_directory)

for file in os.listdir(working_directory):
    if zipfile.is_zipfile(file):
        with zipfile.ZipFile(file) as item:
            item.extractall()
