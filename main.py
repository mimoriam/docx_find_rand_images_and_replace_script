from docx import Document
import os
import glob
import shutil
import random
import re


def get_input_and_check_path():
    while True:
        doc = input("Enter document's path w/ name & file extension: ")
        if os.path.exists(doc):
            return doc
        else:
            print("Either the path or file name is wrong. Please re-input!")


def main():
    doc = get_input_and_check_path()

    document = Document(doc)
    fulltext = []
    username = os.getlogin()

    for p in document.paragraphs:
        fulltext.append(p.text)

    for i in fulltext:
        os.chdir(f'C:\\Users\\{username}\\Desktop\\Images')

        i = re.sub('(.*):\s?', '', i)

        find = glob.glob(f'{os.getcwd()}/**/{i.strip()}.*', recursive=True)

        print(find)

        if not find:
            continue

        findRand = find

        findNotRand = find[0]
        findRandNotList = random.randint(0, (len(findRand) - 1))

        if len(find) == 1:
            findRand = findRand[0]

            shutil.copyfile(
                f'{findNotRand}', f'C:\\Users\\{username}\\Documents\\{os.path.basename(findNotRand)}')
            continue

        shutil.copyfile(f'{find[findRandNotList]}',
                        f'C:\\Users\\{username}\\Documents\\{os.path.basename(find[findRandNotList])}')


if __name__ == '__main__':
    main()
