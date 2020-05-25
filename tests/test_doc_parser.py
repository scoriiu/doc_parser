import json
import os

import pytest
import hashlib

from parser.parser import parse_nb_patients, convert_pdf_to_txt, match_area_of_interest, parse_study_year_range

script_dir = os.path.dirname(os.path.realpath(__file__))


@pytest.fixture(scope='module')
def reference_text():
    text = open(f'{script_dir}/data/text.json')
    return json.load(text)


def test_patients_nb(reference_text):
    for k, v in reference_text.items():
        nb_patients = parse_nb_patients(v['abstract'])[1]
        assert nb_patients, v['patients']


@pytest.mark.parametrize("pdf_filename, m_hash, loc_range", [
    (
            'doc1.pdf',
            '45efc5d85f3c824adce6cfaf166870fb3a5628bbc42f362eef5ba03399efa311',
            (1843, 10664)
    ),
    (
            'doc2.pdf',
            'd052409c48757e6f211b659e33b60f49d15c9576227cdd1181c78d3514688f87',
            (2184, 13738)
    ),
    (
            'doc3.pdf',
            '6d8d51c97fe0e6ec9add63da7f394f6ff1510c8f912eb40c6ed4d2d38120a286',
            (3011, 15710)
    ),
])
def test_match_area_of_interest(pdf_filename, m_hash, loc_range):
    text = convert_pdf_to_txt(f'{script_dir}/data/{pdf_filename}')
    text = text.replace('\n', '').replace('\r', '')
    text_matched, is_matched, loc = match_area_of_interest(text)

    assert is_matched
    assert loc == loc_range
    assert hashlib.sha256(text_matched.encode('utf-8')).hexdigest() == m_hash


def test_study_date(reference_text):
    for k, v in reference_text.items():
        period_of_study = parse_study_year_range(v['materials'], v['abstract'])[0]
        assert period_of_study == tuple(v['period_of_study'])
