# Esxcel

Simple script to create an Excel report on Elasticsearch data

Contribute!

## Install

    git clone https://github.com/MisoMadao/esxcel
    cd esxcel
    pipenv install

## Usage

Please take a look at the [wiki](https://github.com/MisoMadao/esxcel/wiki)!

    usage: es2exc - Elasticsearch query to Excel [-h] [--version] {cli,conf} ...

    Query Elasticsearch and create an excel report with the result

    positional arguments:
      {cli,conf}
        cli       Command line arguments
        conf      Configuration file

    optional arguments:
      -h, --help  show this help message and exit
      --version   show program's version number and exit

For configuration file usage

    usage: es2exc - Elasticsearch query to Excel conf [-h] [--conf CONF]

    optional arguments:
      -h, --help   show this help message and exit
      --conf CONF  path to condfiguration file

For command line usage

    usage: es2exc - Elasticsearch query to Excel cli [-h] [--host HOST] --index
                                                     INDEX --query QUERY
                                                     [--output OUTPUT]
                                                     [--piechart] [--barchart]
                                                     [--user USER]
                                                     [--password PASSWORD]

    optional arguments:
      -h, --help           show this help message and exit
      --host HOST          host:port to make the query
      --index INDEX        index (pattern) to make the query
      --query QUERY        es query to make, every aggregation will be a table
      --output OUTPUT      output file name
      --piechart           add a pie chart from aggregations
      --barchart           add a bar chart from aggregations
      --user USER          username for elasticsearch
      --password PASSWORD  password for elasticsearch

## Result

![](https://image.ibb.co/cwLWEz/001.png)
![](https://image.ibb.co/g7vEZz/003.png)
