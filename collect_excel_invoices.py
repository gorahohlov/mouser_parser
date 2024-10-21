import os
import re
import subprocess
import pandas as pd


pd.set_option('future.no_silent_downcasting', True)


PATH = r'//147.45.104.168/share/Adelante_БР/Запросы от Вертикали'
# folder_name = r'/140-HYVD-195 CNY'
# file_name= r'/commercial invoice HYVD-195 CNY.xlsx'
expected_colheader = ['Description',
                      'Part No.',
                      'HS Code',
                      'COO',
                      'Quantity(pieces)',
                      'Unit price(CNY)',
                      'Brand',
                      'Total (CNY)']
pattern1 = re.compile(r'^\d+-.*', re.I)
pattern2 = re.compile(r'\.xls.*$', re.I)
pattern3 = re.compile(r'commercial invoice .*\.xls.*$', re.I)
invoices_df = pd.DataFrame()
processed_dirs = []
rejected_dirs = []


def find_header_row(path, header_list):
    excel_data = pd.ExcelFile(path)
    for sheet_name in excel_data.sheet_names:
        sheet = excel_data.parse(sheet_name, header=None)
        if sheet.isnull().all(1).sum() == len(sheet):
            continue
        for idx, row in sheet[10:21].iterrows():
            if set(header_list).issubset(set(row.dropna())):
                print(idx, sheet_name)
                return idx, sheet_name
    return None, None
#     raise ValueError('Header row not found')
#     for no, sheet_name in enumerate(excel_data.sheet_names, start=1):
#         print(f'{no}) {sheet_name}')
#     reply = input(f'Введите число от 1 до {len(excel_data.sheet_names)}'
#                   f', для пропуска нажмите Enter: ')
#     print()
#     try:
#         no = int(reply) - 1
#     except ValueError:
#         return '', ''
#     else:
#         if no >= 0 and no < len(excel_data.sheet_names):
#             return , excel_data.sheet_names[no]
#         else:
#             return '', ''


def read_table(path):
    header_row_no, sheet_name = find_header_row(path, expected_colheader)
    if header_row_no and sheet_name:
        global invoices_df
        df = pd.read_excel(path,
                           sheet_name=sheet_name,
                           header=header_row_no,
                           usecols=range(8),
                           names=['DSCR', 'PARTNO', 'HSCODE', 'COO',
                                  'QUAN', 'PRICE', 'BRAND', 'TOTAL'],
                           dtype={'HSCODE': str}
                           )
        df = df.iloc[:df[df.isnull().all(1)].index[0]]
        df['QUAN'] = df['QUAN'].replace('[\$,]', '', regex=True).astype(int)
        invoices_df = pd.concat([invoices_df, df], ignore_index=True)
        return sheet_name, len(df)
    else:
        return None, None

if __name__ == '__main__':

    dir_list = [dir_name for dir_name in os.listdir(PATH) if
                bool(pattern1.search(dir_name))]

    for counter, dir in enumerate(dir_list):
        files_list = os.listdir(os.path.join(PATH, dir))
        target_file = ''
        flag_1 = False
        for each_file in files_list:
            if pattern3.match(each_file):
                target_file = os.path.join(PATH, dir, each_file)
                flag_1 = True
                break
#         if not target_file:
#             xls_files_list = [file for file in files_list if
#                                           bool(pattern2.search(file))]
#             print(f'\nТекущая директория: {dir}')
#             for offset, item in enumerate(xls_files_list, start=1):
#                 print(f'{offset}. {item}')
#             reply = input(f'Введите число от 1 до {len(xls_files_list)}'
#                         f', для пропуска введите 0 или Enter: ')
#             print()
#             try:
#                 int(reply)
#             except ValueError:
#                 target_file = ''
#             else:
#                 num = int(reply) - 1
#                 if num >= 0 and num < len(xls_files_list):
#                     target_file = xls_files_list[num]
#                 else:
#                     target_file = ''
        if target_file:
            path = os.path.join(PATH, dir, target_file)
            sheet_name, lines_number = read_table(path)
            if sheet_name and lines_number:
                processed_dirs.append(dir)
                print(f'Файл: "{dir}/{each_file if flag_1 else target_file}" '
                      f'добавлен;\nЛист: "{sheet_name}", '
                      f'Число строк: "{lines_number}"\n')
            else: rejected_dirs.append((dir, os.path.join(PATH, dir)))
        else:
            rejected_dirs.append((dir, os.path.join(PATH, dir)))
            continue

    processed_dirs_df = pd.DataFrame({'processed_dirs': processed_dirs})
    rejected_dirs_df = pd.DataFrame(rejected_dirs, columns=[
                                                        'наименование_папки',
                                                        'путь_папки'])
    with pd.ExcelWriter(r'c:/users/user/dev/mouser_parse_proj/invoices.xlsx', mode='a',
                        if_sheet_exists='replace') as writer:
        invoices_df.to_excel(writer, sheet_name='invoices', na_rep='Nan')
        processed_dirs_df.to_excel(writer, sheet_name='processed_dirs')
        rejected_dirs_df.to_excel(writer, sheet_name='rejected_dirs')
    print('Обработка инвойсов завершилась')




