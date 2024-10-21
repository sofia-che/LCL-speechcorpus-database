import docx
import pandas as pd
import os

docx_files = [f for f in os.listdir('G:/.shortcut-targets-by-id/1octp96ff7k5QQ7A3nrkuVrrJucEYMO6s/для девочек (комп. лингв)/оригиналы ОРД/') if f.endswith('.docx')]

for docx_file in docx_files:
    print(f"Processing {docx_file}...")
    df_temp = pd.DataFrame(columns=['Time', 'Speaker', 'Text'])

    doc = docx.Document(f'G:/.shortcut-targets-by-id/1octp96ff7k5QQ7A3nrkuVrrJucEYMO6s/для девочек (комп. лингв)/оригиналы ОРД/{docx_file}')
    text = []
    for paragraph in doc.paragraphs:
        text.append(paragraph.text)

    for i in text[2:]:
        timecode = i.strip('[').split('] ')[0]
        newtext = i.strip('[').split('] ')[1:]
        for j in newtext:
            sp = j.split(': ')[0]
            r = j.split(': ')[1].encode('utf-8').decode('utf-8')

            row = {'Time': timecode, 'Speaker': sp, 'Text': r}
            df_temp = df_temp.append(row, ignore_index=True)

    df_temp.to_excel(f'{docx_file.split(".docx")[0]}.xlsx', index=False)

    print(f"{docx_file} processed successfully!")

print("Done!")