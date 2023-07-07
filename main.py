from docx import Document
import os

def read_table(path):
    doc = Document(path)
    table = None
    try:
        table = doc.tables[0]
    except:
        print("File contains no table")
        return None
    
    table_data = []
    header = None
    for i, data in enumerate(table.rows):
        text = (cell.text for cell in data.cells)

        if i==0:
            header = tuple(text)
            continue
        
        table_data.append(
            dict(
                zip(
                    header,
                    text
                )
            )
        )
    return table_data

def search_table(table, keyword):
    found_list = []
    for row in table:
        if any(keyword in string for string in list(row.values())):
            found_list.append(row)
    return found_list

print("Enter keyword: ")
key = input()
print("Enter target directory: ")
directory = input()

for filename in os.listdir(directory):
    if '.docx' not in filename:
        continue
    f = os.path.join(directory, filename)

    print(search_table(read_table(f), key))
