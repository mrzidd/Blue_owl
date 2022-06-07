import zipfile
from lxml import etree
import pandas as pd
import os
import re
#get current working directory
cwd = os.getcwd()

FILE_PATH = cwd

file__ = input("Enter the file name: ")

class Application():
    def __init__(self):
        #debug print('Initialized!')
        xml_content = self.get_word_xml(r'{}/{}'.format(FILE_PATH,file__)) 
        xml_tree = self.get_xml_tree(xml_content)

    def get_word_xml(self, docx_filename):
        with open(docx_filename, 'rb') as f:
            zip = zipfile.ZipFile(f)
            xml_content = zip.read('word/document.xml')
            zip.extractall(FILE_PATH)
        # read the content of txt file in 
        


        return xml_content
    def read_txt(self,file_path):
        """
        Reads the text file and returns a list of strings
        """
        with open(file_path, 'r') as f:
            lines = f.readlines()
        return lines

    def get_xml_tree(self, xml_string):
        return (etree.fromstring(xml_string))


if __name__ == '__main__':
    app = Application()
    data= app.read_txt(f'{FILE_PATH}/word/_rels/document.xml.rels')
    #print(data)
    # find Target="http in data regex
    found = re.findall(r'Target="(.*?)"', str(data))
    print(found)
    # find http and html in found
    found = re.findall(r'http://|https://|http:|https:|html', str(found))
    print(found)
    if len(found) > 0:
        print('your file has been tampered with or has been infected with a virus')


