# This code is designed to clean and reorganize data from an Excel file.
# It reads the data, processes it to remove unnecessary whitespace, and rearranges certain columns.

import pandas as pd
import re
import argparse
import os

def clean_text(text):
    return re.sub(r'\s+', '', str(text)).replace('　', '')

def clean_excel(filename, output_filename):

    # 自動建立 output 目錄
    output_dir = os.path.dirname(output_filename)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    df = pd.read_excel(filename, sheet_name=0, header=None)

    cols_per_record = 6
    num_records = df.shape[1] // cols_per_record

    all_records = []

    for _, row in df.iterrows():
        for i in range(num_records):
            record = row[i * cols_per_record : (i + 1) * cols_per_record].tolist()

            if all((pd.isna(x) or x == '') for x in record):
                continue

            if pd.notna(record[4]):
                record[4] = clean_text(record[4])

            all_records.append(record)

    for idx, rec in enumerate(all_records):
        if (idx + 1) % 10 == 0:
            rec[3], rec[4], rec[5] = rec[2], rec[3], rec[4]
            rec[4] = clean_text(rec[4])
            rec[2] = ''

    # 最後一步：整批複製 D→F
    for rec in all_records:
        rec[5] = rec[3]

    df_out = pd.DataFrame(all_records, columns=['A', 'B', 'C', 'D', 'E', 'F'])
    df_out.to_excel(output_filename, index=False)

    print(f'✅ 已輸出 {output_filename}')

def main():
    parser = argparse.ArgumentParser(description='Excel Cleaner 工具')
    parser.add_argument('input_file', help='輸入 Excel 檔名')
    parser.add_argument('output_file', help='輸出 Excel 檔名')

    args = parser.parse_args()

    clean_excel(args.input_file, args.output_file)

if __name__ == '__main__':
    main()
