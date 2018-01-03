# Esxcel

Simple script to create an Excel report on Elasticsearch data

## Install

    git clone https://github.com/MisoMadao/esxcel
    cd esxcel
    pipenv install

## Usage

    usage: es2exc - Elasticsearch query to Excel [-h] [--version] [--host HOST]
                                                 --index INDEX --query QUERY
                                                 [--output OUTPUT] [--piechart]
                                                 [--barchart]

    Query Elasticsearch and create an excel report with the result

    optional arguments:
      -h, --help       show this help message and exit
      --version        show program's version number and exit
      --host HOST      host:port to make the query (default: 127.0.0.1:9200)
      --index INDEX    index (pattern) to make the query (default: None)
      --query QUERY    es query to make, every aggregation will be a table
                       (default: None)
      --output OUTPUT  output file name (default: es2exc_output.xlsx)
      --piechart       add a pie chart from aggregations (default: False)
      --barchart       add a bar chart from aggregations (default: False)
