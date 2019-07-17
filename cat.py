import ETL
from simple_salesforce import Salesforce
from distutils.util import strtobool
import sys
import argparse

parser = argparse.ArgumentParser(description='generate cat data into salesforce')

parser.add_argument('-i', '--input',
                    help=' cat excel path',
                    required=True)
parser.add_argument('-u', '--username',
                    help=' Username',
                    required=False)

parser.add_argument('-p', '--password',
                    help=' Password',
                    required=False)

parser.add_argument('-t', '--token',
                    help=' Token',
                    required=False,
                    default='')

parser.add_argument('-s', '--sandbox',
                    type= strtobool,
                    help=' is sandbox',
                    required=False,
                    default=True)

parser.add_argument('-v', '--verbose',
                    help=' debug information',
                    required=False,
                    default=False,
                    action='store_true'
                    )


args = parser.parse_args(sys.argv[1:])
etl = ETL.ETL(args.input)

sf = None
if args.username:
    sf = Salesforce(username=args.username, password=args.password, sandbox=args.sandbox, security_token=args.token)
etl.transfer(sf, args)