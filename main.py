#!/usr/bin/env python3
#        1         2         3         4         5         6         7
# 34567890123456789012345678901234567890123456789012345678901234567890123456789
import subprocess
import platform
import time
import os
from typing import Any, Dict, List, Optional
from dotenv import load_dotenv

import pandas as pd
import requests
from pprint3x import pprint

load_dotenv()

# EXCEL_PATH = r'c:/users/user/dev/mouser_parse_proj/mouser_articles.xlsx'
# FILE_PATH = 'commercial invoice LTTC-011.xlsx'
FILE_PATH = input('Введите имя файла из "c:/users/user/downloads/": ')
EXCEL_PATH = r'c:/users/user/downloads/' + FILE_PATH
SHEET_NAME = 'articles'
API_KEY = os.getenv('MOUSER_API_KEY')
d = {'SearchByKeywordRequest': {'keyword': '', }}
url = f'https://api.mouser.com/api/v2.0/search/keyword?apiKey={API_KEY}'

article_frame = pd.DataFrame()
article_series = pd.Series()
article_frame_old = pd.DataFrame()
part_frame_old = pd.DataFrame()
compliance_frame_old = pd.DataFrame()
attributes_frame_old = pd.DataFrame()
pricebreak_frame_old = pd.DataFrame()
errors_frame_old = pd.DataFrame()

attrs_set = set(['Article', 'Description', 'article_no',
                 'ManufacturerPartNumber', 'Manufacturer', 'MouserPartNumber',
                 'Category', 'DataSheetUrl', 'ProductDetailUrl', 'ImagePath',
                 'part_no'])
parts_list: List[Dict] = []
compliance_list: List[Dict] = []
attributes_list: List[Dict] = []
pricebreak_list: List[Dict] = []
article_list: List[Dict] = []
errors_list: List[Dict] = []
article_dict: Dict[str, Any] = {}
flag1: bool = True  # show server response details (by each request)
flag2: bool = True  # show part's details (by each response)
flag4: bool = True  # MaxCallPerDay error flag; if False don't send request
art_no: int = 0


def write_sheets_toexcel(**kwargs):
    """Write processing dataframes to excel file."""
    dataframes: Dict[str, pd.DataFrame] = {}
    for name, frames in kwargs.items():
        df = pd.DataFrame(frames[1],
                          index=range(len(frames[0]) + 1,
                                      len(frames[0]) +
                                      len(frames[1]) + 1)
                          )
        dataframes[name] = pd.concat([frames[0], df])

    with pd.ExcelWriter(EXCEL_PATH,
                        mode='a',
                        if_sheet_exists='replace') as writer:
        dataframes['part_frame'].to_excel(writer,
                                          sheet_name='parts',
                                          na_rep='NaN')
        dataframes['compliance_frame'].to_excel(writer,
                                                sheet_name='compliance')
        dataframes['attributes_frame'].to_excel(writer,
                                                sheet_name='attributes')
        dataframes['pricebreak_frame'].to_excel(writer,
                                                sheet_name='pricebreak')
        dataframes['article_frame'].to_excel(writer,
                                             sheet_name='articles')
        dataframes['errors_frame'].to_excel(writer,
                                            sheet_name='errors')


def x_string_gen(articles: int, num: int, length: int = 48) -> str:
    """Generate status bar string."""
#     articles = 1000 if articles > 1000 else articles
    len_x_str = int(num // (articles / length))
    x_string = 'x' * len_x_str + '.' * (length - len_x_str)
    return f'processing status: [{x_string}]'


def clear_screen() -> None:
    if platform.system() == "Windows":
        subprocess.run("cls", shell=True)
    else:
        subprocess.run("clear")


def output(articles: int, num: int, current_article: str, resp) -> None:
    """Generate screen outputs."""
#     subprocess.run('clear')
    clear_screen()
    print(x_string_gen(articles, num))
#     print(f'num: {num}')
    print(f'num/total_num: {num}/{articles}')
    print(f'art_no/articles: {art_no}/{total_articles}')
    print(f'article: {current_article}')
    if resp['SearchResults']:
        print(f'NumberOfResult: '
              f'{response.json()["SearchResults"]["NumberOfResult"]}\n')
        print('Error has occured: No\n')
    else:
        print(f'NumberOfResult: {None}\n')
        print('Error has occured: Yes\n')
        print('Errors:')
        pprint(resp['Errors'])
        print()


if __name__ == '__main__':

    try:
        excel_data = pd.ExcelFile(EXCEL_PATH)
    except (ValueError, FileNotFoundError):
        print(f'Такого файла: {EXCEL_PATH} не существует!')
#         pass
    else:
        article_frame = pd.read_excel(EXCEL_PATH,
                                      sheet_name=SHEET_NAME,
                                      dtype={'No': 'Int64',
                                             'article_no': 'float16',
                                             'part_no': 'float16',
                                             'USHTS': 'object',
                                             'CNHTS': 'object',
                                             'CAHTS': 'object',
                                             'JPHTS': 'object',
                                             'KRHTS': 'object',
                                             'TARIC': 'object',
                                             'MXHTS': 'object',
                                             'BRHTS': 'object',
                                             'ECCN': 'object'}
                                      )
        article_frame['Article'] = article_frame['Article'].str.strip()
        if attrs_set.issubset(article_frame.columns):
            article_frame.set_index('No', inplace=True)
            article_frame_old = article_frame[~article_frame['article_no'
                                                             ].isna()]
            art_no = len(article_frame_old)
            article_series = article_frame[article_frame['article_no'
                                                         ].isna()]['Article']
            for sheet_name in excel_data.sheet_names:
                if sheet_name == 'parts':
                    part_frame_old = pd.read_excel(EXCEL_PATH,
                                                   sheet_name='parts',
                                                   dtype={
                                                       'No': 'Int64',
                                                       'FactoryStock':
                                                       'float32',
                                                       'Min': 'float32',
                                                       'Mult': 'float32',
                                                       'MultiSimBlue':
                                                       'float32',
                                                       'AvailabilityInStock':
                                                       'float32',
                                                       'article_no': 'int16',
                                                       'part_no': 'int16',
                                                       'SalesMaximumOrderQty':
                                                       'float32'}
                                                   ).set_index('No')
                elif sheet_name == 'compliance':
                    compliance_frame_old = pd.read_excel(
                                                  EXCEL_PATH,
                                                  sheet_name='compliance',
                                                  dtype={
                                                      'No': 'Int64',
                                                      'article_no': 'int16',
                                                      'part_no': 'int16',
                                                      'USHTS': 'object',
                                                      'CNHTS': 'object',
                                                      'CAHTS': 'object',
                                                      'JPHTS': 'object',
                                                      'KRHTS': 'object',
                                                      'TARIC': 'object',
                                                      'MXHTS': 'object',
                                                      'BRHTS': 'object',
                                                      'ECCN': 'object'}
                                                         ).set_index('No')
                elif sheet_name == 'attributes':
                    attributes_frame_old = pd.read_excel(
                                                     EXCEL_PATH,
                                                     sheet_name='attributes',
                                                     dtype={
                                                         'No': 'Int64',
                                                         'article_no': 'int16',
                                                         'part_no': 'int16'}
                                                         ).set_index('No')
                elif sheet_name == 'pricebreak':
                    pricebreak_frame_old = pd.read_excel(
                                                EXCEL_PATH,
                                                sheet_name='pricebreak',
                                                dtype={
                                                       'No': 'Int64',
                                                       'Quantity': 'int32',
                                                       'article_no': 'int16',
                                                       'part_no': 'int16'}
                                                         ).set_index('No')
                    pricebreak_frame_old['Price'] = pricebreak_frame_old[
                            'Price'].replace(r'[\$,]',
                                             '',
                                             regex=True
                                             ).astype('float32')
                elif sheet_name == 'errors':
                    try:
                        errors_frame_old = pd.read_excel(
                                                 EXCEL_PATH,
                                                 sheet_name='errors',
                                                 dtype={
                                                        'No': 'Int64',
                                                        'article_no': 'int16',
                                                        'Id': 'int16'}
                                                         ).set_index('No')
                    except KeyError:
                        pass
#                       print(f'Лист {sheet_name} пуст!')
        else:
            article_series = article_frame['Article']

    articles: int = len(article_series)
    total_articles: int = len(article_frame_old) + articles
    response: Optional[requests.Response] = None
    for num, each_article in enumerate(article_series, start=1):
        # if num == 48: break
        each_article = str(each_article)
        art_no += 1

        """If MaxCallPerDay error occurs, we won't send remaining articles
        for requests."""
        if flag4:
            d['SearchByKeywordRequest']['keyword'] = each_article
            response = requests.request('post', url, json=d)
#         if not response.ok: continue
        if (response.json()['Errors'] and
                response.json()['Errors'][0]['ResourceKey'
                                             ] == 'MaxCallPerDay'):
            flag4 = False
            output(articles, num, each_article, response.json())
            print('Превышен лимит обращений к серверу в день\n' +
                  'Полученные данные будут сохранены в excel;\n' +
                  'остальные артикулы могут быть обработаны завтра.\n' +
                  'подождите еще немного...')
#             for each_error in response.json()['Errors']:
#                 each_error['Article'] = each_article
#                 each_error['art_no'] = art_no
#                 errors_list.append(each_error)
            article_dict = {}
            article_dict['Article'] = each_article
            article_dict['Description'] = 'MaxCallPerDay'
#             article_dict['article_no'] = art_no
            article_list.append(article_dict)
            continue

        """If MaxCallPerMinute error occurs, we'll generate delay and repeat
        the requests until the response becomes success."""
        delay = 1
        err_append_flag = True  # to avoid adding same error_info many times
        while (response.json()['Errors'] and
               response.json()['Errors'][0][
                                       'ResourceKey'] == 'MaxCallPerMinute'):
            output(articles, num, each_article, response.json())
            print(f'Превышен лимит обращений к серверу в минуту\n'
                  f'Генерируем искусственную задержку {delay} секунд '
                  f'и повторим запрос...\n'
                  f'Ждите ...')
            if err_append_flag:
                for each_error in response.json()['Errors']:
                    each_error['Article'] = each_article
                    each_error['article_no'] = art_no
                    errors_list.append(each_error)
                err_append_flag = False
            time.sleep(delay)
            delay += 1
            response = requests.request('post', url, json=d)
        output(articles, num, each_article, response.json())

        """If the another error occurs, we'll handle this case."""
        if response.json()['Errors']:
            for each_error in response.json()['Errors']:
                each_error['Article'] = each_article
                each_error['article_no'] = art_no
                errors_list.append(each_error)
            article_dict = {}
            article_dict['Article'] = each_article
            article_dict['Description'] = 'InvalidCharacters'
            article_dict['article_no'] = art_no
            article_list.append(article_dict)

        if flag1:
            if input('нажмите любую клавишу для продолжения, ' +
                     'для прекращения подтверждения \"q\": ') == 'q':
                flag1 = False

        part_no: int = 0
        if not response.json()['Errors']:
            results_nums = response.json()["SearchResults"]["NumberOfResult"]

            """If ['SearchResults']['Parts'] is not empty we'll processing
            each part in server response."""
            if not response.json()['SearchResults']['NumberOfResult']:
                article_dict = {}
                article_dict['Article'] = each_article
                article_dict['Description'
                             ] = 'NoSearchResultsFound(EmptyPartList)'
                article_dict['article_no'] = art_no
                article_list.append(article_dict)
            else:
                parts = response.json()['SearchResults']['Parts']
                flag3: bool = True  # Add article_info into article_list
                flag5: bool = True  # if parts does not contain each_article, 
                                    # append first part

                for part in parts:
                    part_no += 1
                    if flag2:
                        # subprocess.run('clear')
                        clear_screen()
                        print(x_string_gen(articles, num))
                        print(f'num/total_num: {num}/{articles}')
                        print(f'art_no/articles: {art_no}/{total_articles}')
                        print(f'article: {each_article}')
                        print(f'NumberOfResult: {results_nums}\n')
                        print('part:')
                        pprint(part)
                        print()
                        if input('нажмите любую клавишу для продолжения, ' +
                                 'для прекращения подтверждения ' +
                                 '\"q\": ') == 'q':
                            flag2 = False

                    """Processing attributes list."""
                    for each_attribute in part['ProductAttributes']:
                        each_attribute['MouserPartNubmer'] = part[
                                                    'MouserPartNumber']
                        each_attribute['ManufacturerPartNumber'] = part[
                                                    'ManufacturerPartNumber']
                        each_attribute['Category'] = part['Category']
                        each_attribute['Description'] = part['Description']
                        each_attribute['article_no'] = art_no
                        each_attribute['part_no'] = part_no
                        attributes_list.append(each_attribute)
                    del part['ProductAttributes']

                    """Processing compliance list."""
                    ProductCompliance: dict = {}
                    ProductCompliance['MouserPartNumber'] = part[
                                                       'MouserPartNumber']
                    ProductCompliance['ManufacturerPartNumber'] = part[
                                                 'ManufacturerPartNumber']
                    ProductCompliance['Category'] = part['Category']
                    ProductCompliance['Description'] = part['Description']
                    ProductCompliance['article_no'] = art_no
                    ProductCompliance['part_no'] = part_no
                    for each_compliance in part['ProductCompliance']:
                        ProductCompliance[each_compliance[
                                'ComplianceName']] = each_compliance[
                                'ComplianceValue']
                    compliance_list.append(ProductCompliance)
                    del part['ProductCompliance']

                    """Processing pricelist block."""
                    for each_price in part['PriceBreaks']:
                        each_price['MouserPartNumber'] = part[
                                                    'MouserPartNumber']
                        each_price['ManufacturerPartNumber'] = part[
                                                    'ManufacturerPartNumber']
                        each_price['Category'] = part['Category']
                        each_price['Description'] = part['Description']
                        each_price['article_no'] = art_no
                        each_price['part_no'] = part_no
                        pricebreak_list.append(each_price)
                    del part['PriceBreaks']

                    """Processing part list block."""
                    if not part['AvailabilityOnOrder']:
                        part['AvailabilityOnOrder'] = None
                    if not part['InfoMessages']:
                        part['InfoMessages'] = None
                    part['article_no'] = art_no
                    part['part_no'] = part_no
                    parts_list.append(part)

                    """Processing article_sheet."""
                    if flag5:  # if parts does not contain each_article, append first part
                        article_dict_0 = {}
                        article_dict_0['Article'] = each_article
                        article_dict_0['ManufacturerPartNumber'] = part[
                                                    'ManufacturerPartNumber']
                        article_dict_0['Manufacturer'] = part['Manufacturer']
                        article_dict_0['MouserPartNumber'] = part[
                                                    'MouserPartNumber']
                        article_dict_0['Category'] = part['Category']
                        article_dict_0['Description'] = part['Description']
                        article_dict_0['DataSheetUrl'] = part['DataSheetUrl']
                        article_dict_0['ProductDetailUrl'] = part[
                                                        'ProductDetailUrl']
                        article_dict_0['ImagePath'] = part['ImagePath']
                        article_dict_0['article_no'] = art_no
                        article_dict_0['part_no'] = part_no
                        article_dict_0.update(ProductCompliance)
                        flag5 = False
                    if (part['ManufacturerPartNumber'] == each_article and
                            flag3):
                        article_dict = {}
                        article_dict['Article'] = each_article
                        article_dict['ManufacturerPartNumber'] = part[
                                                    'ManufacturerPartNumber']
                        article_dict['Manufacturer'] = part['Manufacturer']
                        article_dict['MouserPartNumber'] = part[
                                                    'MouserPartNumber']
                        article_dict['Category'] = part['Category']
                        article_dict['Description'] = part['Description']
                        article_dict['DataSheetUrl'] = part['DataSheetUrl']
                        article_dict['ProductDetailUrl'] = part[
                                                        'ProductDetailUrl']
                        article_dict['ImagePath'] = part['ImagePath']
                        article_dict['article_no'] = art_no
                        article_dict['part_no'] = part_no
                        article_dict.update(ProductCompliance)
                        flag3 = False
                        article_list.append(article_dict)
                if flag3:
                    # article_dict = {}
                    # article_dict['Article'] = each_article
                    # article_dict['Description'
                                 # ] = 'ThereAreNoMatchesWithPartList'
                    # article_dict['article_no'] = art_no
                    article_list.append(article_dict_0)

    write_sheets_toexcel(**{'part_frame': (part_frame_old, parts_list),
                            'compliance_frame': (compliance_frame_old,
                                                 compliance_list),
                            'attributes_frame': (attributes_frame_old,
                                                 attributes_list),
                            'pricebreak_frame': (pricebreak_frame_old,
                                                 pricebreak_list),
                            'article_frame': (article_frame_old,
                                              article_list),
                            'errors_frame': (errors_frame_old, errors_list)
                            }
                         )

#     part_frame = pd.DataFrame(parts_list,
#                               index=range(len(part_frame_old) + 1,
#                                           len(part_frame_old) +
#                                           len(parts_list) + 1))
#     part_frame = pd.concat([part_frame_old, part_frame])
#     compliance_frame = pd.DataFrame(compliance_list,
#                                    index=range(len(compliance_frame_old) + 1,
#                                                len(compliance_frame_old) +
#                                                len(compliance_list) + 1))
#     compliance_frame = pd.concat([compliance_frame_old, compliance_frame])
#     attributes_frame = pd.DataFrame(attributes_list,
#                                    index=range(len(attributes_frame_old) + 1,
#                                                len(attributes_frame_old) +
#                                                len(attributes_list) + 1))
#     attributes_frame = pd.concat([attributes_frame_old, attributes_frame])
#     pricebreak_frame = pd.DataFrame(pricebreak_list,
#                                    index=range(len(pricebreak_frame_old) + 1,
#                                                len(pricebreak_frame_old) +
#                                                len(pricebreak_list) + 1))
#     pricebreak_frame = pd.concat([pricebreak_frame_old, pricebreak_frame])
#     article_frame = pd.DataFrame(article_list,
#                                  index=range(len(article_frame_old) + 1,
#                                              total_articles + 1))
#     article_frame = pd.concat([article_frame_old, article_frame])
#     errors_frame = pd.DataFrame(errors_list,
#                                 index=range(len(errors_frame_old) + 1,
#                                             len(errors_frame_old) +
#                                             len(errors_list) + 1))
#     errors_frame = pd.concat([errors_frame_old, errors_frame])
#
#     with pd.ExcelWriter(EXCEL_PATH, mode='a', if_sheet_exists='replace') \
#             as writer:
#         part_frame.to_excel(writer, sheet_name='parts', na_rep='NaN')
#         compliance_frame.to_excel(writer, sheet_name='compliance')
#         attributes_frame.to_excel(writer, sheet_name='attributes')
#         pricebreak_frame.to_excel(writer, sheet_name='pricebreak')
#         article_frame.to_excel(writer, sheet_name='articles')
#         errors_frame.to_excel(writer, sheet_name='errors')
    print('Data download from mouser server has been successfully completed')
    print('Загрузка данных с сервера mouser.com была успешно выполнена')
