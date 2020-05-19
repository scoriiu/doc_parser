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

introduction_title_pattern = f'{regex_ignore_case("introduction")}[1-9\\s]*[A-Z]'
methods_title_pattern = f'({regex_ignore_case("methods")}|{regex_ignore_case("patients")}|{regex_ignore_case("materials")})[1-9\\s]*[A-Z]'
discussion_title_pattern = f'{regex_ignore_case("discussion")}[1-9\\s]*[A-Z]'
references_title_pattern = f'{regex_ignore_case("references")}[1-9\\s]*[A-Z]'
pattern_nb_patients = \
    r'(.{50})(\d[\d,]*)([^\d%]{0,50}((?i)patients|cases|subjects|individuals))(.{30})'


def convert_pdf_to_txt(path) -> str:
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ''
    maxpages = 0
    caching = True
    pagenos = set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching,
                                  check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text


def log_processing_done(filename):
    log = f'{filename} '
    while len(log) < 110:
        log += '.'
    log += ' Done.'
    print(log)


def compute_search_pattern(text):
    max_ix = 1e7
    references_ix = re.search(references_title_pattern, text, re.IGNORECASE)
    references_ix = references_ix.regs[0] if references_ix else max_ix
    references_ix = references_ix[0] if type(references_ix) == tuple else references_ix

    discussion_ix = re.search(discussion_title_pattern, text, re.IGNORECASE)
    discussion_ix = discussion_ix.regs[0] if discussion_ix else max_ix
    discussion_ix = discussion_ix[0] if type(discussion_ix) == tuple else discussion_ix

    end_pattern = references_title_pattern if references_ix < discussion_ix else discussion_title_pattern
    return f'(?={methods_title_pattern})(.*?)(?={end_pattern})'


def parse_nb_patients(text) -> (int, str):
    m = re.findall(pattern_nb_patients, text[:2000], re.IGNORECASE)
    if not m:
        return 0, ''

    max_group = max(m, key=lambda g: int(g[1].replace(',', '')))
    return max_group[1], f'{max_group[1]}\n({"".join(max_group[:3]) + max_group[-1]})'


def compute_results(pdf_dir, keywords):
    results = {}
    pdf_files = glob.glob(f'{pdf_dir}/*.pdf')
    for pdf_file in pdf_files:
        pdf_name = pdf_file.split('/')[-1] if '/' in pdf_file else pdf_file
        pdf_name = pdf_file.split('\\')[-1] if '\\' in pdf_file else pdf_name
        text = convert_pdf_to_txt(pdf_file)
        text = text.replace('\n', '').replace('\r', '')

        pattern = compute_search_pattern(text)
        text_constrained = text
        m = re.findall(pattern, text)
        if not m:
            area_of_interest_matched = False
        else:
            area_of_interest_matched = True
            text_constrained = m[0][1]

        keywords_search_results = {'#Patients': parse_nb_patients(text)[1]}
        for k_name, v_keywords in keywords.items():
            _text = text if k_name in global_keywords else text_constrained
            flags = 0 if k_name == 'Country' else re.IGNORECASE
            matches = [k for k in v_keywords if re.search(f'\\b{k}\\b', _text, flags)]
            keywords_search_results[k_name] = matches

        keywords_search_results['AreaOfInterestMatched'] = area_of_interest_matched
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

    headers = ['Document', '#Patients', *keywords, ]
    for ix, header in enumerate(headers):
        worksheet.write(0, ix, header, bold_style)

    cells = []
    for k, v in results.items():
        doc_name = k
        doc_name += ' (*)' if not v['AreaOfInterestMatched'] else ''
        row = [doc_name]
        for _k, _v in v.items():
            if _k not in [*keywords.keys(), '#Patients']:
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
    headers = ['<b>Document<b>', '<b>#Patients<b>'] + [f'<b>{k}<b>' for k in keywords]
    cells = []
    for k, v in results.items():
        doc_name = k
        doc_name += ' (*)' if not v['AreaOfInterestMatched'] else ''
        row = [doc_name]
        for _k, _v in v.items():
            if _k not in [*keywords.keys(), '#Patients']:
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
