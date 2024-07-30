# litmos-chris21-integration

This PowerShell script creates/updates Litmos users from a Chris21 XML export.

If you find this script helpful, please consider supporting ACMI buy making a donation via https://tickets.acmi.net.au/donate/i/23

------------------------
Installation/Setup
------------------------

1. Place the PS script in your desired location
2. Insure that Chris21 is generating an XML export file in same location as the script
3. Update $importFile so that it references the name of your Chris21 export
4. Update $apiKey with your Litmos API key
5. Function 'Build-XMLBody' will need to be modified to match your Litmos elements and element order
6. 'Main' will also need to be modified as above
7. Manager email logic can be omitted if not needed 
8. Run the PS script manually or via Window Task Scheduler as required

------------------------
Notes
------------------------
1. The script will check and create/re-create its required file structure
2. Logs are produced for each run (success or failure)
3. Error logs are produced for each error
4. Supplementary logs show XML body being passed to API
5. Chris21 exports are archived on each run
6. Archived exports are deleted if older than 60 days
