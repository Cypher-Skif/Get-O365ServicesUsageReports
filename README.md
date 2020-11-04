# Get-O365UsageDetails
Get report about O365 services usage, can generate summary O365 services usage reports and send it through on-premise exchange relay.
Available reports:   
- Office365 Active User Detail  
- SharePoint Site Usage Detail  
- SkypeForBusiness Activity User Detail  
- Teams User Activity User Detail  
- OneDrive Usage Account Detail  

Use GenerateSecret.ps1 to ecnrypt ClientSecret  
Use Settings.xml to configure script for your tenant.  
![Settings.xml](https://github.com/Cypher-Skif/PublicRepoPictures/blob/master/Get-O365ServicesUsageReports_Settings.png)  
Required API Permissions:
- Reports.Read.All

