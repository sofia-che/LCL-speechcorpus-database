import pandas as pd
import glob
import os

folder_path = 'C:/Users/София/Downloads/оригиналы ОРД xlsx/'

files = glob.glob(os.path.join(folder_path, '*.xlsx'))

for file in files:
    print(f'Opening file: {file}')
    df = pd.read_excel(file)
    df.insert(0, 'File_name', os.path.basename(file))
    df.insert(df.shape[1], 'Speaker_2', '')
    df.insert(df.shape[1], 'Comments', '')
    df.insert(df.shape[1], 'Com_type', '')
    df.to_excel(file, index=False)
    print(f'File {file} has been updated.')

print('All files have been updated.')