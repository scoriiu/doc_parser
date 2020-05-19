import json
import os
import re

import pytest
import hashlib

from parser.parser import parse_nb_patients, convert_pdf_to_txt, compute_search_pattern

script_dir = os.path.dirname(os.path.realpath(__file__))


@pytest.fixture(scope='module')
def text():
    text = open(f'{script_dir}/data/text.json')
    return json.load(text)


def test_patients_nb(text):
    for k, v in text.items():
        nb_patients = parse_nb_patients(v['abstract'])[1]
        assert nb_patients, v['patients']


@pytest.mark.parametrize("pdf_filename, m_hash", [
    ('test_area_of_interest1.pdf', 'd8f80cae2e3066ac22adac64c2d27a6ab7931b5177cef329627323152954ea11'),
    ('test_area_of_interest2.pdf', 'b53c86847f7ff2eca60498fe2d8dd75d2042fb4cf69fafe61857734928c3e5a8'),
])
def test_match_area_of_interest(pdf_filename, m_hash):
    text = convert_pdf_to_txt(f'{script_dir}/data/{pdf_filename}')
    text = text.replace('\n', '').replace('\r', '')
    pattern = compute_search_pattern(text[1000:-1000])

    assert pattern == '(?=([Mm][Ee][Tt][Hh][Oo][Dd][Ss]|[Pp][Aa][Tt][Ii][Ee][Nn][Tt][Ss]' \
                      '|[Mm][Aa][Tt][Ee][Rr][Ii][Aa][Ll][Ss])[1-9\\s]*[A-Z])(.*?)' \
                      '(?=[Dd][Ii][Ss][Cc][Uu][Ss][Ss][Ii][Oo][Nn][1-9\\s]*[A-Z])'

    m = re.findall(pattern, text)
    assert hashlib.sha256(m[0][1].encode('utf-8')).hexdigest() == m_hash


