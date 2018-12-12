import ETL
from simple_salesforce import Salesforce


etl = ETL.ETL('input/Contract Assessment Tool - Renewal-2017.xlsm')

#etl.generate_object_meta()
#etl.field_map()
#etl.generate_config('fieldmap.json')
etl.transfer(Salesforce(username='', password='', sandbox=True, security_token=''))