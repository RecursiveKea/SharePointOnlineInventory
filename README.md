# SharePoint Online Inventory
Export SharePoint Online Site Collection, Webs, Lists, and Item properties to CSV

### Notes before using
- Script is free to use (and extend) and is provided as-is.
- Credentials being used needs to have access to all the SharePoint sites to inventory
- When the script completes make sure to check that there were no errors
- Throttling is a concern for this script. There are some elements in the script (eg Get-SPOnlineHelperPnPProperty) to help mitigate this however make sure to review the output to see if there are any errors thrown
- If you run this over an entire tenant I would recommend excluding the permissions as it is a massive bottleneck for the script. Threading this helps. Going to look into how to speed this up. Scanning items also takes a while but isn’t quite as time consuming as permissions.
- Change the following if required
  - The Date Format is hardcoded at the moment (dd-MMM-yyyy HH:mm:ss.fff), change this if required.
  - As the dates are returned in UTC so are converted into the time zone of the machine running the script (regardless of what the site is set to). This is set in the inventory settings. You can hard code this to your preferred time zone.
- Detecting if the SharePoint list items have unique permissions is done via the REST API. The reason for this is it’s faster as this property can be retrieved in bulk, PnP has to load this per-item.
- Schema XML most of the time can be excluded unless you have a requirement for it. Having this will vastly increase the size of the CSV file extracts.
- As this script generates CSV files line breaks are replaced. If you need them you can have the function have them replaced with “{LineBreak}” then re-add the line break when processing it for the import (SQL, Power BI, etc).
- Using the “WaitBetweenSites” will force a 1 minute wait between sites to help further prevent throttling.
- Review the excluded system fields (this list is incomplete and there are some system fields I left in (eg Created, Modified)

### Data Model
<img width="1521" alt="Data Model" src="https://user-images.githubusercontent.com/102898289/168424990-66562400-8409-4c34-af66-767b1093969b.png">
