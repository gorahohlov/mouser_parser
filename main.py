#        0         2         3         4         5         6         7
# 34567890123456789012345678901234567890123456789012345678901234567890123456789
import subprocess
import time

import pandas as pd
import requests
from pprint3x import pprint

EXCEL_PATH = r'/mnt/c/users/user/dev/mouser_parse_proj/mouser_articles.xlsx'
SHEET_NAME = 'articles'
API_KEY = 'a6adbd72-cc89-4454-a34c-dd0f75364d52'
d = {'SearchByKeywordRequest': {'keyword': '', }}
url = f'https://api.mouser.com/api/v1.0/search/keyword?apiKey={API_KEY}'

article_frame = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
article_frame['Article'] = article_frame['Article'].str.strip()

parts_list = []
compliance_list = []
attributes_list = []
pricebreak_list = []
article_list = []
errors_list = []
flag1 = True
flag2 = True
flag4 = True
art_no = 0
total_articles = len(article_frame['Article'])


def x_string_gen(total_articles: int, art_no: int, length: int = 48) -> str:
    """Generate status bar string
    """
    len_x_str = int(art_no // (total_articles / length))
    x_string = 'x' * len_x_str + '.' * (length - len_x_str)
    return f'processing status: [{x_string}]'


def output(total_art: int, art_no: int, current_article: str, resp) -> None:
    """Generate screen outputs
    """
    subprocess.run('clear')
    print(x_string_gen(total_art, art_no))
    print(f'art_no: {art_no}')
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


for each_article in article_frame['Article']:
    each_article = str(each_article)
    art_no += 1

    """If MaxCallPerDay error occurs, we won't send remaining articels
    for requests."""
    if flag4:
        d['SearchByKeywordRequest']['keyword'] = each_article
        response = requests.request('post', url, json=d)
    if (response.json()['Errors'] and
            response.json()['Errors'][0]['ResourceKey'] == 'MaxCallPerDay'):
        flag4 = False
        output(total_articles, art_no, each_article, response.json())
        print('Превышен лимит обращений к серверу в день\n' +
              'Полученные данные будут сохранены в excel;\n' +
              'остальные артикулы могут быть обработаны завтра.\n' +
              'подождите еще немного...')
        for each_error in response.json()['Errors']:
            each_error['Article'] = each_article
            each_error['art_no'] = art_no
            errors_list.append(each_error)
        article_dict = {}
        article_dict['Article'] = each_article
        article_dict['Description'] = 'MaxCallPerDay'
        article_dict['article_no'] = art_no
        article_list.append(article_dict)
        continue

    """If MaxCallPerMinute error occurs, we'll generate delay and repeat
    the requests until the response becomes success."""
    delay = 1
    while (response.json()['Errors'] and
           response.json()['Errors'][0]['ResourceKey'] == 'MaxCallPerMinute'):
        output(total_articles, art_no, each_article, response.json())
        print(f'Превышен лимит обращений к серверу в минуту\n'
              f'Генерируем искусственную задержку {delay} секунд '
              f'и повторим запрос...\n'
              f'Ждите ...')
        for each_error in response.json()['Errors']:
            each_error['Article'] = each_article
            each_error['art_no'] = art_no
            errors_list.append(each_error)
        time.sleep(delay)
        delay += 1
        response = requests.request('post', url, json=d)
    output(total_articles, art_no, each_article, response.json())

    """If the another error occurs, we'll handle this case"""
    if response.json()['Errors']:
        for each_error in response.json()['Errors']:
            each_error['Article'] = each_article
            each_error['art_no'] = art_no
            errors_list.append(each_error)
        article_dict = {}
        article_dict['Article'] = each_article
        article_dict['Description'] = 'InvalidCharacters'
        article_dict['article_no'] = art_no
        article_list.append(article_dict)

    if flag1:
        if input('нажмите любую клавишу, для прекращения ' +
                 'подтверждения \"q\": ') == 'q':
            flag1 = False

    part_no = 0
    if not response.json()['Errors']:

        """If ['SearchResults']['Parts'] is not empty we'll processing
        each part in server response"""
        if not response.json()['SearchResults']['NumberOfResult']:
            article_dict = {}
            article_dict['Article'] = each_article
            article_dict['Description'] = 'HasNoSearchResults(EmptyPartList)'
            article_dict['article_no'] = art_no
            article_list.append(article_dict)
        else:
            parts = response.json()['SearchResults']['Parts']
            flag3 = True
            for part in parts:
                part_no += 1
                if flag2:
                    subprocess.run('clear')
                    print(x_string_gen(total_articles, art_no))
                    print(f'art_no: {art_no}')
                    print(f'article: {each_article}')
                    print(f'NumberOfResult: '
                          f'{response.json()["SearchResults"]["NumberOfResult"]}\n')
                    print('part:')
                    pprint(part)
                    print()
                    if input('нажмите любую клавишу, для прекращения ' +
                             'подтверждения \"q\": ') == 'q':
                        flag2 = False

                """processing attributes list"""
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

                """processing compliance list"""
                ProductCompliance = {}
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
                        'ComplianceName']] = each_compliance['ComplianceValue']
                compliance_list.append(ProductCompliance)
                del part['ProductCompliance']

                """processing pricelist block"""
                for each_price in part['PriceBreaks']:
                    each_price['MouserPartNumber'] = part['MouserPartNumber']
                    each_price['ManufacturerPartNumber'] = part[
                                                    'ManufacturerPartNumber']
                    each_price['Category'] = part['Category']
                    each_price['Description'] = part['Description']
                    each_price['article_no'] = art_no
                    each_price['part_no'] = part_no
                    pricebreak_list.append(each_price)
                del part['PriceBreaks']

                """processing part list block"""
                if not part['AvailabilityOnOrder']:
                    part['AvailabilityOnOrder'] = None
                if not part['InfoMessages']: part['InfoMessages'] = None
                part['article_no'] = art_no
                part['part_no'] = part_no
                parts_list.append(part)

                """processing article_sheet"""
                if part['ManufacturerPartNumber'] == each_article and flag3:
                    article_dict = {}
                    article_dict['Article'] = each_article
                    article_dict['ManufacturerPartNumber'] = part[
                                                    'ManufacturerPartNumber']
                    article_dict['Manufacturer'] = part['Manufacturer']
                    article_dict['MouserPartNumber'] = part['MouserPartNumber']
                    article_dict['Category'] = part['Category']
                    article_dict['Description'] = part['Description']
                    article_dict['DataSheetUrl'] = part['DataSheetUrl']
                    article_dict['ProductDetailUrl'] = part['ProductDetailUrl']
                    article_dict['ImagePath'] = part['ImagePath']
                    article_dict['article_no'] = art_no
                    article_dict['part_no'] = part_no
                    article_dict.update(ProductCompliance)
                    flag3 = False
                    article_list.append(article_dict)
            if flag3:
                article_dict = {}
                article_dict['Article'] = each_article
                article_dict['Description'] = 'ThereIsNoMatchesWithPartList'
                article_dict['article_no'] = art_no
                article_list.append(article_dict)

part_frame = pd.DataFrame(parts_list,
                          index=range(1, len(parts_list) + 1))
compliance_frame = pd.DataFrame(compliance_list,
                                index=range(1, len(compliance_list) + 1))
attributes_frame = pd.DataFrame(attributes_list,
                                index=range(1, len(attributes_list) + 1))
pricebreak_frame = pd.DataFrame(pricebreak_list,
                                index=range(1, len(pricebreak_list) + 1))
article_frame = pd.DataFrame(article_list,
                             index=range(1, len(article_list) + 1))
errors_frame = pd.DataFrame(errors_list,
                            index=range(1, len(errors_list) + 1))

with pd.ExcelWriter(EXCEL_PATH, mode='a', if_sheet_exists='replace') as writer:
    part_frame.to_excel(writer, sheet_name='parts', na_rep='NaN')
    compliance_frame.to_excel(writer, sheet_name='compliance')
    attributes_frame.to_excel(writer, sheet_name='attributes')
    pricebreak_frame.to_excel(writer, sheet_name='pricebreak')
    article_frame.to_excel(writer, sheet_name='articles')
    errors_frame.to_excel(writer, sheet_name='errors')
print('Data download from mouser server has been successfully completed')
print('Загрузка данных с сервера mouser.com была успешно выполнена')
