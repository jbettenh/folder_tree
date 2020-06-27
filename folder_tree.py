#! python 3
# This will list out all folders and subfolders for a given directory.

import argparse
import os
import docx
from docx.enum.dml import MSO_THEME_COLOR_INDEX


def dir_path(string):
    if os.path.isdir(string):
        return string
    else:
        raise NotADirectoryError(string)


def arguments():
    parser = argparse.ArgumentParser(description='Parser')
    parser.add_argument('--path', type=dir_path, help='This is the path to the folder')

    return parser.parse_args()

def add_hyperlink(paragraph, text, url):
    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )

    # Create a w:r element and a new w:rPr element
    new_run = docx.oxml.shared.OxmlElement('w:r')
    rPr = docx.oxml.shared.OxmlElement('w:rPr')

    # Join all the xml elements together add add the required text to the w:r element
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    # Create a new Run object and add the hyperlink into it
    r = paragraph.add_run ()
    r._r.append (hyperlink)

    # A workaround for the lack of a hyperlink style (doesn't go purple after using the link)
    # Delete this if using a template that has the hyperlink style in it
    r.font.color.theme_color = MSO_THEME_COLOR_INDEX.HYPERLINK
    r.font.underline = True

    return hyperlink


def main():
    args = arguments()

    maindir = (os.path.basename(args.path))
    print(maindir)

    for dirpath, dirnames, files in os.walk(args.path):
        path = dirpath.split('/')
        if dirnames:
            for folder in dirnames:
                print('-' + folder)

    document = docx.Document()
    p = document.add_paragraph('Folder: ')
    add_hyperlink(p, 'folder name', dirpath)
    document.save('test/demo_hyperlink.docx')

if __name__ == '__main__':
    main()