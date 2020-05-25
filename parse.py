import argparse
import json
import os
import sys

import glob
import json
import re

import plotly.graph_objects as go

import xlsxwriter

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO

from utils.utils import regex_ignore_case

global_keywords = ['Country']

introduction_title_pattern = f'{regex_ignore_case("introduction")}[1-9\\s]*[1-9A-Z]'
area_of_interest_start_pattern = f'({regex_ignore_case("methods")}|{regex_ignore_case("patients")}|{regex_ignore_case("materials")}*)[1-9\\s]*[1-9A-Z]'
area_of_interest_end_pattern = f'({regex_ignore_case("discussion")}|{regex_ignore_case("references")})[1-9\\s]*[1-9A-Z]'
study_date_pattern = r'(.{0,20})((19|20)\d{2})(.{0,25})((19|20)\d{2})(.{0,20})'
pattern_nb_patients = \
    r'(.{50})(\d[\d,]*)([^\d%]{0,50}((?i)patients|cases|subjects|individuals))(.{30})'


def convert_pdf_to_txt(path) -> str:
    rsrc_mgr = PDFResourceManager()
    ret_str = StringIO()
    la_params = LAParams()
    device = TextConverter(rsrc_mgr, ret_str, laparams=la_params)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrc_mgr, device)
    password = ''
    max_pages = 0
    caching = True
    page_nos = set()

    for page in PDFPage.get_pages(fp, page_nos, maxpages=max_pages, password=password,
                                  caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = ret_str.getvalue()

    fp.close()
    device.close()
    ret_str.close()
    return text


def log_processing_done(filename):
    log = f'{filename} '
    while len(log) < 110:
        log += '.'
    log += ' Done.'
    print(log)


def match_area_of_interest(text) -> (str, bool, (int, int)):
    start_loc = re.search(area_of_interest_start_pattern, text)
    start_loc = min(start_loc.regs, key=lambda l: l[0]) if start_loc else 1000
    start_loc = start_loc[0] if type(start_loc) == tuple else start_loc

    end_loc = re.search(area_of_interest_end_pattern, text)
    end_loc = min(end_loc.regs, key=lambda l: l[1]) if end_loc else len(text) - 1000
    end_loc = end_loc[1] if type(end_loc) == tuple else end_loc

    is_match = start_loc != 1000 and end_loc != len(text) - 1000
    return text[start_loc:end_loc], is_match, (start_loc, end_loc)


def parse_nb_patients(text) -> (int, str):
    m = re.findall(pattern_nb_patients, text[:2000], re.IGNORECASE)
    if not m:
        return 0, ''

    max_group = max(m, key=lambda g: int(g[1].replace(',', '')))
    return max_group[1], f'{max_group[1]}\n({"".join(max_group[:3]) + max_group[-1]})'


def parse_study_year_range(text, optional_text=None) -> ((int, int), str):
    m = re.findall(study_date_pattern, text, re.IGNORECASE)
    if not m and optional_text:
        m = re.findall(study_date_pattern, optional_text, re.IGNORECASE)

    if not m:
        return (0, 0), ''

    year_range = int(m[0][1]), int(m[0][4])
    return year_range, f'{m[0][1]}-{m[0][4]}\n({m[0][0] + m[0][1] + m[0][3] + m[0][4] + m[0][6]})'


def compute_results(pdf_dir, keywords):
    results = {}
    pdf_files = glob.glob(f'{pdf_dir}/*.pdf')
    for pdf_file in pdf_files:
        pdf_name = pdf_file.split('/')[-1] if '/' in pdf_file else pdf_file
        pdf_name = pdf_file.split('\\')[-1] if '\\' in pdf_file else pdf_name
        text = convert_pdf_to_txt(pdf_file)
        text = text.replace('\n', '').replace('\r', '')

        text_of_interest, text_of_interest_matched, loc_range = match_area_of_interest(text)
        study_year_range = parse_study_year_range(text[loc_range[0]:loc_range[0] + 1000], text[:2000])[1]
        keywords_search_results = {
            '#Patients': parse_nb_patients(text)[1],
            'Period Of Study': study_year_range
        }
        for k_name, v_keywords in keywords.items():
            _text = text if k_name in global_keywords else text_of_interest
            flags = 0 if k_name == 'Country' else re.IGNORECASE
            matches = [k for k in v_keywords if re.search(f'\\b{k}\\b', _text, flags)]
            keywords_search_results[k_name] = matches

        keywords_search_results['AreaOfInterestMatched'] = text_of_interest_matched
        results[pdf_name] = keywords_search_results

        log_processing_done(pdf_name)

    return results


def log_results(results):
    print('\nResults:')
    for k_pdf, v_matches in results.items():
        print('-' * 100)
        print(f'Document: {k_pdf}')
        print(json.dumps(v_matches, indent=4))


def export_to_excel(results, keywords, output):
    basic_format = {'text_wrap': 40, 'text_h_align': 2}
    header_format = {'bold': True, 'bg_color': 'silver', **basic_format}
    warn_format = {'font_color': 'red', **basic_format}

    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    basic_style = workbook.add_format(basic_format)
    bold_style = workbook.add_format(header_format)
    warn_style = workbook.add_format(warn_format)

    headers = ['Document', '#Patients', 'Period Of Study', *keywords, ]
    for ix, header in enumerate(headers):
        worksheet.write(0, ix, header, bold_style)

    cells = []
    for k, v in results.items():
        doc_name = k
        doc_name += ' (*)' if not v['AreaOfInterestMatched'] else ''
        row = [doc_name]
        for _k, _v in v.items():
            if _k not in [*keywords.keys(), '#Patients', 'Period Of Study']:
                continue

            cell = ', '.join(_v) if isinstance(_v, list) else _v
            row.append(cell)

        cells.append(row)

    worksheet.set_column(0, len(headers), 40)
    for row_ix, row in enumerate(cells):
        for coll_ix, cell in enumerate(row):
            worksheet.write(row_ix + 1, coll_ix, cell, warn_style if '(*)' in cell else basic_style)

    workbook.close()


def export_to_html(results, keywords):
    header_color = 'grey'
    row_even_color = 'lightgrey'
    row_odd_color = 'white'
    headers = ['<b>Document<b>', '<b>#Patients<b>', '<b>Period Of Study<b>'] + [f'<b>{k}<b>' for k in keywords]
    cells = []
    for k, v in results.items():
        doc_name = k
        doc_name += ' (*)' if not v['AreaOfInterestMatched'] else ''
        row = [doc_name]
        for _k, _v in v.items():
            if _k not in [*keywords.keys(), '#Patients', 'Period Of Study']:
                continue

            cell = ', '.join(_v) if isinstance(_v, list) else _v
            row.append(cell)

        cells.append(row)

    fig = go.Figure(data=[go.Table(
        header=dict(
            values=headers,
            line_color='darkslategray',
            fill_color=header_color,
            align=['left', 'center'],
            font=dict(color='white', size=12)
        ),
        cells=dict(
            height=80,
            values=list(zip(*cells)),
            line_color='darkslategray',
            fill_color=[[row_odd_color, row_even_color, row_odd_color, row_even_color, row_odd_color] * len(headers)],
            align=['left', 'center'],
            font=dict(color='darkslategray', size=11)
        ))
    ])

    fig.show()


def export_to_json(results, output):
    with open(output, 'w') as outfile:
        json.dump(results, outfile, indent=4)


print(sys.path)

def parse_args():
    parser = argparse.ArgumentParser(description='Parses documents based on multiple criteria')
    parser.add_argument(
        '--pdf_dir',
        help='Path to the dir containing the documents',
        required=True,
    )

    parser.add_argument(
        '--html',
        help='Exports the results to html',
        default=False,
        action='store_true',
        required=False,
    )

    parser.add_argument(
        '--excel',
        help='Exports the results to excel',
        default=False,
        action='store_true',
        required=False,
    )

    parser.add_argument(
        '--test',
        help='Unit testing',
        default=False,
        action='store_true',
        required=False,
    )

    parser.add_argument(
        '--verbose',
        default=False,
        action='store_true',
        required=False,
    )

    return parser.parse_args()


if __name__ == '__main__':
    script_dir = os.path.dirname(os.path.realpath(__file__))
    results_dir = f'{script_dir}/results'
    os.makedirs(script_dir, exist_ok=True)
    keywords = open(f'{script_dir}/keywords.json')
    keywords = json.load(keywords)

    args = parse_args()
    results = compute_results(args.pdf_dir, keywords)

    pdf_dir = args.pdf_dir[:-1] if args.pdf_dir[-1] in ['\\', '/'] else args.pdf_dir
    export_to_json(results, os.path.join(results_dir, f'results_{pdf_dir}.json'))

    if args.verbose:
        log_results(results)

    if args.excel:
        export_to_excel(results, keywords, os.path.join(results_dir, f'results_{pdf_dir}.xlsx'))

    if args.html:
        export_to_html(results, keywords)
