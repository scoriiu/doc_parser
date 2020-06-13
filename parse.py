import argparse
import json
import os

from parser.parser import compute_results, export_to_excel, export_to_html, log_results, export_to_json


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
    results_dir = os.path.join(script_dir, 'results')
    if not os.path.exists(results_dir):
        os.makedirs(results_dir)

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
