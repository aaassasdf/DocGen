import os
from docxtpl import *
import jinja2
import csv
from docxcompose.composer import Composer
from docx import Document

# Read points.csv to gather infomations
def read_csv(BASE_DIR: str):
    items = []
    desc = []
    info = dict()
    with open(rf'{BASE_DIR}\points.csv',"r", encoding='utf-8-sig',errors='ignore') as file:
        csvfile = csv.reader(file)
        for row in csvfile:
            if row[0].isalpha():
                info[row[0]] = row[1]
            else:
                items.append(row[0])
                desc.append(row[1])

    return items, desc, info

def create_doc_context(BASE_DIR: str):
    frameworks = [] # The structure of the whole document
    framework = [] # The structure of each single document
    items, desc, info = read_csv(BASE_DIR)

    for i,d in zip(items,desc):
        #scan contents line-by-line
        
        framework.append({'item':i,'desc':d, 'result':'Pass □ / Fail □ / NA □'})

        # Once it reaches to the last page, padding empty string to framework until 10 page contents are created
        # to make sure all the tables' size accross all pages are the same
        if int(i) == len(items):
            for _ in range(10-len(framework)):
                framework.append({'item':'','desc':'','result':''})

        # Once scanned 10 items or the last items, append to the resulting document holder and reset the temporary content holder to empty
        if int(i)%10 == 0 or int(i) == len(items):
            frameworks.append(framework)
            framework = [] 

    # resulting frameworks structure is [[page_1_content][page_2_content]...]
    context = {'frameworks':frameworks}
    return frameworks, info

def combine_docs(BASE_DIR:str, pages:int):
    master = Document(rf'{BASE_DIR}\T&C forms\T&C form0.docx')
    composer = Composer(master)
    for page in range(pages+1):

        if os.path.exists(rf'{BASE_DIR}\T&C forms\T&C form{page}.docx'):
            temp_doc = Document(rf'{BASE_DIR}\T&C forms\T&C form{page}.docx')
            composer.append(temp_doc)
            os.remove(rf'{BASE_DIR}\T&C forms\T&C form{page}.docx')
        else:
            print(f'Error occured: File T&C form{page} does not exist')
            break

    composer.save(rf'{BASE_DIR}\T&C forms\T&C form.docx')


def create_doc(BASE_DIR: str):

    doc = DocxTemplate(rf'{BASE_DIR}\T&C template.docx')
    frameworks, info= create_doc_context(BASE_DIR)

    for page,framework in enumerate(frameworks):
        context = {'frameworks':framework}
        context = context | info
        jinja_env = jinja2.Environment(autoescape=True)
        doc.render(context, jinja_env)
        doc.save(rf'{BASE_DIR}\T&C forms\T&C form{page}.docx')

    combine_docs(BASE_DIR, page)

if __name__ == "__main__":
    print("processing...")
    BASE_DIR = os.path.abspath('auto-gen Document')
    create_doc(BASE_DIR)
    print("--Completed--")
    print(f'Document is created in {BASE_DIR}\\T&C forms')