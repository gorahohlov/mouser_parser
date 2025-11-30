import subprocess
import platform
import time
from typing import Any, Dict, List, Optional

import requests
from pprint3x import pprint as pp

API_KEY = 'a6adbd72-cc89-4454-a34c-dd0f75364d52'
d = {'SearchByKeywordRequest': {'keyword': '', }}
url = f'https://api.mouser.com/api/v2.0/search/keyword?apiKey={API_KEY}'
article_series = [
    'AD8051ARZ-REEL7',
    '09031966921',
    '09185106323',
    '09692815421',
    '09670254715',
    '09670009922',
    '09670254715',
    '09670094715',
    '0900BL15C050E',
    '502585-0470',
    '51065-0500',
    'M80-0030006',
    'VS-10CTQ150-M3',
    'LT5538IDD#TRPBF',
    'DLFCV-1600+',
    'GRM31CR71E106KA12L',
    'GRM32ER71H106KA12L',
    'BZV55-B10,115',
    '09062486823',
    'DW-10-09-G-D-300',
    'APHCM2012SYCK-F01',
    'APHCM2012ZGC-F01',
    'WSL2512R0100FEA18',
    'FTSH-113-01-F-MT',
]
parts_list: List[Dict] = []
compliance_list: List[Dict] = []
attributes_list: List[Dict] = []
pricebreak_list: List[Dict] = []
article_list: List[Dict] = []
errors_list: List[Dict] = []
article_dict: Dict[str, Any] = {}
article_dict_0: Dict[str, str] = {}
flag1: bool = True  # show server response details (by each request)
flag2: bool = True  # show part's details (by each response)
flag4: bool = True  # MaxCallPerDay error flag; if False don't send request
art_no: int = 0


def x_string_gen(articles: int, num: int, length: int = 48) -> str:
    """Generate status bar string."""
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
        pp(resp['Errors'])
        print()


def article_info(part, art_no, part_no):
    """Generate dict of article head info of each Respose[Part]"""
    return {
        'MouserPartNumber': part['MouserPartNumber'],
        'ManufacturerPartNumber': part['ManufacturerPartNumber'],
        'Category': part['Category'],
        'Description': part['Description'] ,
        'article_no': art_no,
        'part_no': part_no,
    }

def process_attributes(part, art_no, part_no):
    base = article_info(part, art_no, part_no)
    for attr in part.pop('ProductAttributes', []):
        attributes_list.append({**attr, **base})

def process_compliance(part, art_no, part_no):
    base = article_info(part, art_no, part_no)
    compliance = {**base}
    for comp in part.pop('ProductCompliance', []):
        compliance[comp['ComplianceName']] = comp['ComplianceValue']
    compliance_list.append(compliance)
    return compliance

def process_pricebreaks(part, art_no, part_no):
    base = article_info(part, art_no, part_no)
    for price in part.pop('PriceBreaks', []):
        pricebreak_list.append({**price, **base})

def process_part(part, art_no, part_no):
    part['AvailabilityOnOrder']= part.get('AvailabilityOnOrder') or None
    part['InfoMessages'] = part.get('InfoMessages') or None
    part.update(article_info(part, art_no, part_no))
    parts_list.append(part)

def make_article_dict(part, each_article, art_no, part_no, compliance):
    return {
            'Article': each_article,
            'ManufacturerPartNumber': part['ManufacturerPartNumber'],
            'Manufacturer': part['Manufacturer'],
            'MouserPartNumber': part['MouserPartNumber'],
            'Category': part['Category'],
            'Description': part['Description'],
            'DataSheetUrl': part['DataSheetUrl'],
            'ProductDetailUrl': part['ProductDetailUrl'],
            'ImagePath': part['ImagePath'],
            'article_no': art_no,
            'part_no': part_no,
            **compliance,
    }


if __name__ == '__main__':

    articles: int = len(article_series)
    total_articles: int = len(article_series) 
    response: Optional[requests.Response] = None
    for num, each_article in enumerate(article_series, start=1):
        each_article = str(each_article)
        art_no += 1

        """If MaxCallPerDay error occurs, we won't send remaining articles
        for requests."""
        if flag4:
            d['SearchByKeywordRequest']['keyword'] = each_article
            response = requests.post(url, json=d)
        if (response.json()['Errors'] and
                response.json()['Errors'][0]['ResourceKey'
                                             ] == 'MaxCallPerDay'):
            flag4 = False
            output(articles, num, each_article, response.json())
            print('Превышен лимит обращений к серверу в день\n' +
                  'Полученные данные будут сохранены в excel;\n' +
                  'остальные артикулы могут быть обработаны завтра.\n' +
                  'подождите еще немного...')
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
                # flag5 is used as a marker:
                # if parts does not contain each_article,
                # then append the first part in parts;
                flag5: bool = True
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
                        pp(part)
                        print()
                        if input('нажмите любую клавишу для продолжения, ' +
                                 'для прекращения подтверждения ' +
                                 '\"q\": ') == 'q':
                            flag2 = False

                    """Processing attributes list."""
                    process_attributes(part, art_no, part_no)

                    """Processing compliance list."""
                    compliance = process_compliance(part, art_no, part_no)

                    """Processing pricelist block."""
                    process_pricebreaks(part, art_no, part_no)

                    """Processing part list block."""
                    process_part(part, art_no, part_no)

                    """Processing article_sheet."""
                    if flag5:
                        article_dict_0 = make_article_dict(
                            part,
                            each_article,
                            art_no,
                            part_no,
                            compliance,
                        )
                        flag5 = False
                    if (
                        part['ManufacturerPartNumber'] == each_article 
                        and flag3
                    ):
                        article_list.append(
                            make_article_dict(
                                part,
                                each_article,
                                art_no,
                                part_no,
                                compliance
                            )
                        )
                        flag3 = False
                if flag3:
                    article_list.append(article_dict_0)
