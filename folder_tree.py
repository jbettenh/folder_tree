#! python 3
# This will list out all folders and subfolders for a given directory.

import argparse
from pathlib import Path
import docx
from docx.enum.dml import MSO_THEME_COLOR_INDEX

# prefix components:
space =  '    '
branch = '│   '
# pointers:
tee =    '├── '
last =   '└── '


def tree(dir_path: Path, prefix: str=''):
    """A recursive generator, given a directory Path object
    will yield a visual tree structure line by line
    with each line prefixed by the same characters
    """
    contents = list(dir_path.iterdir())
    # contents each get pointers that are ├── with a final └── :
    pointers = [tee] * (len(contents) - 1) + [last]
    for pointer, path in zip(pointers, contents):
        yield prefix + pointer + path.name
        if path.is_dir(): # extend the prefix and recurse:
            extension = branch if pointer == tee else space
            # i.e. space because last, └── , above so no more |
            yield from tree(path, prefix=prefix+extension)


def arguments():
    parser = argparse.ArgumentParser(description='Parser')
    parser.add_argument('--path', type=Path,default=Path(__file__).absolute().parent / "Users", help='This is the path to the folder')

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

    for line in tree(args.path/ 'Documents'):
        print(line)

   # for dirpath, dirnames, files in os.walk(args.path):
    #    path = dirpath.split('/')
     #   if dirnames:
      #      for folder in dirnames:
       #         print('-' + folder)

    document = docx.Document()
    p = document.add_paragraph('Folder: ')
    add_hyperlink(p, 'folder name', dirpath)
    document.save('test/demo_hyperlink.docx')

if __name__ == '__main__':
    main()