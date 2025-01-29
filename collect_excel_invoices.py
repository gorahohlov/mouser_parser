import os
import re
from typing import List

import openpyxl
import pandas as pd


pd.set_option('future.no_silent_downcasting', True)


PATH = r'//147.45.104.168/share/Adelante_БР/Запросы от Вертикали/'
path_to_save_xlsx_file = r'c:/users/user/dev/mouser_parse_proj/'
expected_colheader = ['Description',
                      'Part No.',
                      'HS Code',
                      'COO',
                      'Quantity(pieces)',
                      'Unit price(CNY)',
                      'Brand',
                      'Total (CNY)']
pattern_11 = re.compile(r'^\d+-.*', re.I)
pattern_12 = re.compile(r'\.xls.*$', re.I)
pattern_13 = re.compile(r'commercial invoice .*\.xls.*$', re.I)
invoices_df = pd.DataFrame()
inv_header_df = pd.DataFrame()
header_attrs_list: List[str] = ['Invoice number',
                                'Date',
                                'shipper',
                                'Contact',
                                'Receiver:',
                                'Contract:',
                                'Contract date:']
inv_header_list = []
processed_dirs = []
rejected_dirs = []


def get_inv_header_attrs(header_series):
    inv_header_dict = {}
    for attr in header_attrs_list:
        try:
            i = header_series.index.get_loc(header_series[header_series ==
                                                          attr].index[0]) + 1
        except KeyError:
            inv_header_dict[attr] = None
        else:
            inv_header_dict[attr] = header_series.iloc[i]
    return inv_header_dict


def find_header_row(path, header_list):
    try:
        excel_data = pd.ExcelFile(path)
    except ValueError:
        return None, None, None
    else:
        for sheet_name in excel_data.sheet_names:
            sheet = excel_data.parse(sheet_name, header=None)
            if sheet.isnull().all(1).sum() == len(sheet):
                continue
            for idx, row in sheet[10:21].iterrows():
                if set(header_list[:4]).issubset(set(row.dropna()[:4])):
                    print(idx, sheet_name)
                    break
            else:
                idx, sheet_name, inv_header_dict = None, None, None
            if idx and sheet_name:
                sheet_header = sheet.iloc[:idx, :8].stack()
                inv_header_dict = get_inv_header_attrs(sheet_header)
                break
        else:
            idx, sheet_name, inv_header_dict = None, None, None
    return idx, sheet_name, inv_header_dict
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
    global invoices_df
    (header_row_no,
     sheet_name,
     inv_header_dict) = find_header_row(path, expected_colheader)
    if header_row_no and sheet_name:
        df = pd.read_excel(path,
                           sheet_name=sheet_name,
                           header=header_row_no,
                           usecols=range(8),
                           names=['DSCR', 'PARTNO', 'HSCODE', 'COO',
                                  'QUAN', 'PRICE', 'BRAND', 'TOTAL'],
                           dtype={'HSCODE': str}
                           )
        df = df.iloc[:df[df.isnull().all(1)].index[0]]
        df['QUAN'] = df['QUAN'].replace(r'[\$,]', '', regex=True).astype(int)
        df = df.assign(**inv_header_dict)
        invoices_df = pd.concat([invoices_df, df], ignore_index=True)
        inv_header_dict['DIR_NAME'] = dir
        inv_header_dict['DIR_PATH'] = path
        inv_header_list.append(inv_header_dict)
        return sheet_name, len(df)
    else:
        return None, None


def select_xls_file(files_list):
    global frag_1
    xls_files_list = [file for file in files_list if
                      bool(pattern_12.search(file))]
    while xls_files_list:
        print(f'\nТекущая директория: {dir}')
        for offset, item in enumerate(xls_files_list, start=1):
            print(f'{offset}. {item}')
        print('0. Отмена (не выбирать никакой файл)')
        reply = input(f'Введите число от 0 до {len(xls_files_list)}'
                      f', для пропуска введите 0: ')
        print()

        try:
            int(reply)
        except ValueError:
            print(f'ValueError! Требуется ввод цифры: '
                  f'от 0..{len(xls_files_list)}. '
                  f'Попробуем еще раз.\n')
            continue
        else:
            num = int(reply) - 1
            if num >= 0 and num < len(xls_files_list):
                file = xls_files_list.pop(num)
                path = os.path.join(PATH, dir, file)
                resp1, resp2 = read_table(path)
                if resp1 and resp2:
                    return resp1, resp2, file
                else:
                    print(f'Этот файл: "{file}" не подходит.\n'
                          f'Он содержит неподходящие данные.\n'
                          f'Давайте попробуем выбрать другой файл.\n')
                    continue
            elif num == -1:
                frag_1 = True
                return None, None, None
            else:
                print(f'Требуется ввод цифры '
                      f'от 0..{len(xls_files_list)}! '
                      f'Попробуем еще раз.\n')
                continue
    print('Мы исчерпали все варианты xls(x) файлов в директории.\n' +
          'Ни один не подошёл.\n' +
          'Или в указанной директории нет файлов excel вообще.\n')
    frag_1 = True
    return None, None, None


if __name__ == '__main__':

    dir_list = [dir_name for dir_name in os.listdir(PATH) if
                bool(pattern_11.search(dir_name))]

    for counter, dir in enumerate(dir_list):
        # if counter == 16: break
        files_list = os.listdir(os.path.join(PATH, dir))
        target_file = ''
        flag_1 = False # True if file was selected by pattern; False - manually
        for each_file in files_list:
            if pattern_13.match(each_file):
                target_file = os.path.join(PATH, dir, each_file)
                flag_1 = True
                path = os.path.join(PATH, dir, target_file)
                sheet_name, lines_number = read_table(path)
                processed_dirs.append(dir)
                print(f'Файл: "{dir}/{each_file if flag_1 else target_file}" '
                      f'добавлен;\nЛист: "{sheet_name}", '
                      f'Число строк: "{lines_number}"\n')
                # break
        # else:
        # if not flag_1:
        flag_2 = False # if valid file was not found in dir; True - was found;
        while not flag_1:
            sheet_name, lines_number, target_file = select_xls_file(files_list)
            if sheet_name and lines_number:
                flag_2 = True
                processed_dirs.append(dir)
                print(f'Файл: "{dir}/{each_file if flag_1 else target_file}" '
                      f'добавлен;\nЛист: "{sheet_name}", '
                      f'Число строк: "{lines_number}"\n')
        if not flag_2:
            rejected_dirs.append((dir, os.path.join(PATH, dir)))
            print(f'В директории: "{dir}" не нашлось подходящего файла для '
                  'импорта данных;\n'
                  'Данные инвойса из этой папки не добавлены.\n')
        continue

    invoices_df.rename(columns={'DSCR': 'DSCR',
                                'PARTNO': 'PARTNO',
                                'HSCODE': 'HSCODE',
                                'COO': 'COO',
                                'QUAN': 'QUAN',
                                'PRICE': 'PRICE',
                                'BRAND': 'BRAND',
                                'TOTAL': 'TOTAL',
                                'Invoice number': 'INVC_NO',
                                'Date': 'INVC_DD',
                                'shipper': 'SHPR',
                                'Contact': 'ATTN',
                                'Receiver:': 'CNEE',
                                'Contract:': 'CONT_NO',
                                'Contract date:': 'CONT_DD'},
                       inplace=True)
    inv_header_df = pd.DataFrame(inv_header_list)
    inv_header_df.rename(columns={'Invoice number': 'INVC_NO',
                                  'Date': 'INVC_DD',
                                  'shipper': 'SHPR',
                                  'Contact': 'ATTN',
                                  'Receiver:': 'CNEE',
                                  'Contract:': 'CONT_NO',
                                  'Contract date:': 'CONT_DD',
                                  'dir_name': 'DIR_NAME',
                                  'dir_path': 'DIR_PATH'
                                  },
                         inplace=True,
                         errors='ignore')
    processed_dirs_df = pd.DataFrame({'processed_dirs': processed_dirs})
    rejected_dirs_df = pd.DataFrame(rejected_dirs, columns=[
                                                        'наименование_папки',
                                                        'путь_папки'])
    if not os.path.isfile(path_to_save_xlsx_file + 'invoices.xlsx'):
        wb = openpyxl.Workbook()
        wb.save(path_to_save_xlsx_file + 'invoices.xlsx')
    with pd.ExcelWriter(path_to_save_xlsx_file + 'invoices.xlsx',
                        mode='a',
                        if_sheet_exists='replace'
                        ) as writer:
        invoices_df.to_excel(writer, sheet_name='invoices', na_rep='Nan')
        inv_header_df.to_excel(writer, sheet_name='inv_headers', na_rep='Nan')
        processed_dirs_df.to_excel(writer, sheet_name='processed_dirs')
        rejected_dirs_df.to_excel(writer, sheet_name='rejected_dirs')
    print('Обработка инвойсов завершилась\n')
    print('Invoice processing has been finished\n')
