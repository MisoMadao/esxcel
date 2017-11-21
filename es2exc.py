import json
import logging

import sys
from elasticsearch import Elasticsearch
import argparse
from openpyxl import Workbook

if __name__ == '__main__':

    argument_parser = argparse.ArgumentParser(
        prog='es2exc - Elasticsearch query to Excel',
        description='Query Elasticsearch and create an excel report with the result',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
        add_help=True
    )
    argument_parser.add_argument('--version', action='version', version='%(prog)s 0.1')
    argument_parser.add_argument('--host', help='host:port to make the query', default='127.0.0.1:9200')
    argument_parser.add_argument('--index', help='index (pattern) to make the query', required=True)
    argument_parser.add_argument('--query', help='es query to make, every aggregation will be a table', required=True)
    argument_parser.add_argument('--output', help='output file name', default='es2exc_output.xlsx')
    args = vars(argument_parser.parse_args())

    e_logger = logging.Logger(__name__)
    e_handler = logging.FileHandler('es2exc.log')
    e_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    e_handler.setFormatter(e_formatter)
    e_logger.addHandler(e_handler)
    e_logger.info(args)

    try:
        args['query'] = args['query'].replace("'", '"')
        es_client = Elasticsearch(hosts=[args['host']])
        args['query'] = json.loads(args['query'])
    except Exception as ex:
        e_logger.exception(ex)
        print('1')
        sys.exit(1)

    es_response = es_client.search(index=args['index'], body=args['query'])

    wb = Workbook()
    ws = wb.get_active_sheet()
    ws.title = 'hits'
    column_next = 1
    column_values = {}
    row_next = 2

    for hit in es_response['hits']['hits']:
        # TODO iter over nested fields
        for k, v in hit['_source'].items():
            if k not in column_values:
                column_values[k] = column_next
                column_next += 1
                ws.cell(row=1, column=column_values[k], value=k)
            ws.cell(row=row_next, column=column_values[k], value=str(v))
        row_next += 1

    for agg_name in es_response['aggregations'].keys():
        column_width = 8
        wb.create_sheet(agg_name)
        ws = wb.get_sheet_by_name(agg_name)
        row = 1
        ws.cell(row=row, column=1, value=agg_name)
        ws.cell(row=row, column=2, value='count')
        for bucket in es_response['aggregations'][agg_name]['buckets']:
            if len(bucket['key']) > column_width:
                column_width = len(bucket['key'])
            row += 1
            ws.cell(row=row, column=1, value=bucket['key'])
            ws.cell(row=row, column=2, value=bucket['doc_count'])
        ws.column_dimensions['A'].width = column_width

    wb.save(args['output'])
