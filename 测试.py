import docx

def extract_data_from_table(doc, keyword1, keyword2):
    result = []
    for table in doc.tables:
        for row in table.rows:
            for i, cell in enumerate(row.cells):
                if keyword1 in cell.text or keyword2 in cell.text:
                    try:
                        next_cell = row.cells[i + 1]
                        result.append((cell.text, next_cell.text))
                    except IndexError:
                        pass
    return result

file_path = "C:/Users/Administrator/Desktop/农村非低保残疾人/测试/20230220-037795薛清乔.docx"
doc = docx.Document(file_path)
keyword1 = "受理人"
keyword2 = "受理时间"

data = extract_data_from_table(doc, keyword1, keyword2)
print(data)
