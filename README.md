# cat-etl
bootstrap process

1. install python 3+
2. install dependencies
               
       pip3 install openpyxl
       pip3 install simple-salesforce
2. create a folder called input and drop an excel renewal cat into it.
4. update the cat.py 

* update file reference to match input file

       etl = ETL.ETL('input/Contract Assessment Tool - Renewal-2017.xlsm')

* put your salesforce credentials into the following line, you'll need username, password and security token

       etl.transfer(Salesforce(username='', password='', sandbox=True, security_token=''))
       
       
* execute the script by calling (Windows/OSX)

         $python3 cat.py
