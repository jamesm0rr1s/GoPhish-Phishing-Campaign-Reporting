# GoPhish Phishing Campaign Reporting (GoPhishReport)

GoPhishReport is a penetration testing and red teaming tool that creates custom phishing campaign reports for GoPhish phishing campaigns based of the two CSV files that GoPhish can generate.

## Main Features

 - Automated phishing campaign reporting
   - A phishing report that combines the GoPhish CSV files
   - A phishing report that summarizes all emails opened, links clicked, and data submitted
 
## Installation

Clone the GitHub repository
```
git clone https://github.com/jamesm0rr1s/GoPhish-Phishing-Campaign-Reporting /opt/jamesm0rr1s/GoPhish-Phishing-Campaign-Reporting
```

## Usage

 - Execute a phishing campaign with GoPhish
 - Export the GoPhish phishing data in both the Events and Results CSV files
 - Update the file names in the PowerShell script (Lines 6 and 7)
 - Update the directory names in the PowerShell script (Lines 10 and 13)
 - Run the following PowerShell script:
```
CreatePhishingReportsFromGoPhishPhishingCampaign.ps1
```

## Example Screenshots

### Input Files

[Example of Events.csv](Example%20Input%20-%20GoPhish%20CSV%20Files/Events.csv)  
![ExampleInput-Events.csv](Example%20Screenshots%20-%20Input%20and%20Output%20Files/Input%20-%20Events.csv.png?raw=true "ExampleInput-Events.csv")

[Example of Results.csv](Example%20Input%20-%20GoPhish%20CSV%20Files/Results.csv)  
![ExampleInput-Results.csv](Example%20Screenshots%20-%20Input%20and%20Output%20Files/Input%20-%20Results.csv.png?raw=true "ExampleInput-Results.csv")

### Output Files

[Example of Phishing Campaign Report 1 - Details](Example%20Output%20-%20Phishing%20Reports/Phishing%20Campaign%20Report%201%20-%20Details.csv)  
![ExampleOutput-PhishingReport1.xlsx](Example%20Screenshots%20-%20Input%20and%20Output%20Files/Output%20-%20Phishing%20Report%201.png?raw=true "ExampleOutput-PhishingReport1.xlsx")

[Example of Phishing Campaign Report 2 - Summary](Example%20Output%20-%20Phishing%20Reports/Phishing%20Campaign%20Report%202%20-%20Summary.csv)  
![ExampleOutput-PhishingReport2.xlsx](Example%20Screenshots%20-%20Input%20and%20Output%20Files/Output%20-%20Phishing%20Report%202.png?raw=true "ExampleOutput-PhishingReport2.xlsx")