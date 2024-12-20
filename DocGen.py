import csv
from pathlib import Path
from docxcompose.composer import Composer
from docx import Document
from docxtpl import DocxTemplate
import jinja2

# actual files to be used
# target_csv = "points.csv"
# target_template = "TnC template.docx"
target_dir = "TnC forms"

#for testing
target_csv = "points_example.csv"
target_template = "TnC template Sample.docx"


# Read points.csv to gather information
def read_csv(base_dir: Path):
    items, desc, info = [], [], {}
    with open(base_dir / target_csv, "r", encoding='utf-8-sig', errors='ignore') as file:
        csvfile = csv.reader(file)

        for row in csvfile:
            if not row[0].isnumeric():
                info[row[0]] = row[1]
            else:
                items.append(row[0])
                desc.append(row[1])

        if not items:
            raise ValueError("Error: No items found in the CSV file.")
        
        if not desc:
            raise ValueError("Error: No descriptions found in the CSV file.")
        
        if not info:
            raise ValueError("Error: No information found in the CSV file.")
        
    return items, desc, info

def create_doc_context(base_dir: Path):
    assert (base_dir / target_csv).exists(), "points.csv does not exist" # Check if the file exists
    frameworks, framework = [], []
    csv_path = base_dir / target_csv
    items, desc, info = read_csv(base_dir)

    for i, d in zip(items, desc):
        i_int = int(i)  # Convert once and reuse
        framework.append({'item': i, 'desc': d, 'result': 'Pass □ / Fail □ / NA □'})

        # Once it reaches the 10th item, append the framework to the frameworks list
        if i_int % 10 == 0:
            frameworks.append(framework)
            framework = []

    # Add padding to the last framework if it's less than 10 and not empty
    if framework:
        padding_needed = 10 - len(framework)
        if padding_needed > 0:
            framework.extend([{'item': '', 'desc': '', 'result': ''} for _ in range(padding_needed)])
        frameworks.append(framework)

    return frameworks, info

def combine_docs(base_dir: Path, num_pages: int):
    assert (base_dir / target_dir).exists(), "TnC forms directory does not exist" # Check if the directory exists
    master_doc_path = base_dir / target_dir / 'TnC form0.docx'
    master = Document(master_doc_path)
    composer = Composer(master)
    
    for page in range(num_pages + 1): # Loop through the pages
        temp_doc_path = base_dir / target_dir / f'TnC form{page}.docx' # Path to the temporary document
        if temp_doc_path.exists(): # Check if the file exists and it's not the master document
            if page > 0:
                temp_doc = Document(temp_doc_path)
                composer.append(temp_doc)
            temp_doc_path.unlink()  # Removes the file
        
        else:
            print(f'Error occurred: File {temp_doc_path} does not exist')
            break

        

    final_doc_path = base_dir / target_dir / 'TnC form.docx'
    composer.save(final_doc_path)

    print(f"Succeed! Document saved to {final_doc_path}")

def create_doc(base_dir: Path):
    assert (base_dir / target_template).exists(), "TnC template.docx does not exist" # Check if the file exists
    template_path = base_dir / target_template
    doc = DocxTemplate(template_path)
    frameworks, info = create_doc_context(base_dir)

    # Check if directory TnC forms exists, if not, create one
    if not (base_dir / target_dir).exists():
        print(f"Directory {target_dir} does not exist. Creating directory...")
        (base_dir / target_dir).mkdir()
        print(f"Directory {target_dir} created.")

    for page, framework in enumerate(frameworks):
        context = {'frameworks': framework} # Create a new dictionary for each page
        context = context | info  # Merge the two dictionaries
        jinja_env = jinja2.Environment(autoescape=True) # Create a Jinja environment
        doc.render(context, jinja_env) # Render the document
        doc.save(base_dir / target_dir / f'TnC form{page}.docx') # Save the document

    combine_docs(base_dir, page) # Combine the documents

if __name__ == "__main__": # Only run the script if it's the main script
    print("processing...")
    base_dir = Path(__file__).parent  # Directory of the script
    create_doc(base_dir)