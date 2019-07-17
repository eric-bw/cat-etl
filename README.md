# cat-etl
bootstrap process

1. install python 3+
2. install dependencies
               
       pip or pip3 install openpyxl
       pip or pip3 install simple-salesforce
2. create a folder called input and drop an excel renewal cat into it.



##Execution       
execute the script by calling (Windows/OSX)

         $python3 cat.py -i ./input/Contract Assesssment -Renewal.xlsm -u <username> -p password -t token -s t -v


## params
        # path to assessment tool (required)
        -i --input (path to file) 
        
        # SF environment credentials (not required)
        -u --username (salesforce username)
        -p --password (salesforce password)
        -t --token (salesforce token)
        -s --sandbox (is environment a sandbox)
        
        # Verbose (not required)
        -v verbose (print extra information about the transfer
        

## verbose output

if you do not provide the salesforce credential information you will get a dry run of the output if you use -v that would be loaded into salesforce.

example:

        ---------------------------------------
        loading sheet:  Update List
        inserting to table:  CAT_Model_UpdateList__c
        record:  {'CAT_Model__c': None, 'Row__c': '1', 'Name': '1', 'A__c': 'Recommended updates'}
        record:  {'CAT_Model__c': None, 'Row__c': '2', 'Name': '2', 'A__c': 'Tony - Refresh data to July - Dec 2016; Also add d'}
        record:  {'CAT_Model__c': None, 'Row__c': '3', 'Name': '3', 'A__c': 'Tony - Add list of hospitals to input tab when "sy'}
        record:  {'CAT_Model__c': None, 'Row__c': '4', 'Name': '4', 'A__c': 'Heather - Add Delivery frequency tier to "historic'}
        record:  {'CAT_Model__c': None, 'Row__c': '5', 'Name': '5', 'A__c': 'Heather - Propose Delivery frequency improvement f'}
        record:  {'CAT_Model__c': None, 'Row__c': '6', 'Name': '6', 'A__c': 'Tony - show surcharges in historical information ('}
        record:  {'CAT_Model__c': None, 'Row__c': '7', 'Name': '7', 'A__c': 'Tony - Combine IRL and Molecular revenue'}
        record:  {'CAT_Model__c': None, 'Row__c': '8', 'Name': '8', 'A__c': 'Tony - Move RDP information into "other revenue" a'}
        record:  {'CAT_Model__c': None, 'Row__c': '9', 'Name': '9', 'A__c': "Heather - remove cells highlighted in red (I didn'"}
        record:  {'CAT_Model__c': None, 'Row__c': '10', 'Name': '10', 'A__c': 'Heather - update costs for IRL and Molecular to be'}
        record:  {'CAT_Model__c': None, 'Row__c': '11', 'Name': '11', 'A__c': 'Heather - can you double-check the mapping of all '}
        record:  {'CAT_Model__c': None, 'Row__c': '12', 'Name': '12', 'A__c': 'Heather - can you add to your macro to hide rows 7'}
        record:  {'CAT_Model__c': None, 'Row__c': '13', 'Name': '13', 'A__c': None}
        record:  {'CAT_Model__c': None, 'Row__c': '14', 'Name': '14', 'A__c': 'Updates for future release'}
        record:  {'CAT_Model__c': None, 'Row__c': '15', 'Name': '15', 'A__c': 'Removal of PSP or inclusion in other revenue?'}
        record:  {'CAT_Model__c': None, 'Row__c': '16', 'Name': '16', 'A__c': 'Addition of PR section to SDP product'}