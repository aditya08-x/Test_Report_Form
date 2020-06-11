import argparse
import json
import smtplib
import traceback
import pandas as pd
import os
from createTestReportPackages.environment_variables import load_environment_variables

CONFIG = None

def start_api(args):
    """ This function loads configurations, dependencies and start python server

    :param args: arguments required to start python server
    """
    print('-----------------------------')
    print('---- Starting Python API ----')
    try:
        args_dict = vars(args)
        host = args_dict['host'][0]
        port = args_dict['port'][0]
        prod = args_dict['prod'][0]
        workers = args_dict['workers'][0]
        threads = args_dict['threads'][0]
        config_file_path = args_dict['configPath'][0]
        print('Host : ', host)
        print('Port : ', port)
        print('Config file path: ', config_file_path)
        json_file = open(config_file_path, 'r')
        global CONFIG
        CONFIG = json.load(json_file)
        CONFIG = load_environment_variables(CONFIG)
        from createTestReportPackages.utils import platform_logging
        platform_logging.set_log_config(CONFIG)
        from createTestReportPackages import server
        server.start_hosting(host=host, port=port, prod=prod,
                             number_of_workers=workers, threads=threads)
    except Exception as exp_obj:
        print('------------------------------------')
        print('---- Failed to Start Python API ----')
        print(f'Exception : {exp_obj}\n{traceback.format_exc()}')
        print('------------------------------------')


def main():
    """ Main entry function """
    try:
        parser = argparse.ArgumentParser()
        subparsers = parser.add_subparsers(help='list of sub parsers')
        api_parser = subparsers.add_parser("Api")
        api_parser.set_defaults(func=start_api)
        api_parser.add_argument("-host", nargs='+', required=False, dest="host", type=str,
                                action='store', default=["localhost"], help="Host IP")
        api_parser.add_argument("-port", nargs='+', required=False, dest="port", type=str,
                                action='store', default=[5000], help="Port number")
        api_parser.add_argument("-configPath", nargs='+', required=True, dest="configPath",
                                type=str, action='store', default=[None], help="config file path")
        api_parser.add_argument("-prod", nargs='+', required=False, dest="prod", type=bool,
                                action='store', default=[False], help="Run Production server")
        api_parser.add_argument("-workers", nargs='+', required=False, dest="workers", type=bool,
                                action='store', default=[3], help="Number of workers")
        api_parser.add_argument("-threads", nargs='+', required=False, dest="threads", type=bool,
                                action='store', default=[3], help="Number of threads")
        args = parser.parse_args()
        args.func(args)
    except Exception as exp_obj:
        print(exp_obj)
