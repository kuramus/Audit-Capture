# Audit Capture
Retrieving the audit from a Dynamics CRM Online instance as a spreadsheet.

## Getting Started
This tool helps in exporting the audit of multiple records queried through a FetchXML from a Dynamics 365 CRM instance, to an Excel spreadsheet.

## Setup
The code can be loaded into Visual Studio. Any dependencies (packages) required can be obtained from NuGet Package Manager.

## Working
The application works in the following process flow.
1. Establish a connection to the CRM instance.
2. Retrieve the records queried using the FetchXML.
3. Capture the Audit for all these records.
4. Export the audit to an excel spreadsheet.

The Audit for different records is retrieved in the excel spreadsheet as multiple rows. Since, the audit retrieved is of multiple records, these records can be separated using the field *Entity Title Attribute*. Please note, this attribute must be included in the rows being retrieved by the FetchXML.

## How to use
1. Paste the FetchXML which returns the records, for which the audit is to be retrieved.
2. Specify the path in the output directory, for the spreadsheet.

The connection with CRM is established using one of the two methods.
#### A. CRM Service Client (New)
The URL or credentials are not required in the fields provided on the first page of the application.

1. Directly click on *Login (CRM Service Client)*.
This opens the login dialog box, which connects to the newer CRM instances and is the recommended approach for connecting with a CRM instance.
You may use 'Show Advanced' to enter the credentials of your Office 365 account.

2. Once logged in, use the *Get Audit (CRM Service Client)* button to fetch the audit and create the spreadsheet.

#### B. CRM Organization Service (Legacy)
This is only for legacy systems where Organization Service is supported and the primary way for establishing connection to the CRM system.

1. Enter the URL for the CRM instance in the application itself along with the credentials.
2. Click on the *Get Audit (Using Org Service)* button to fetch the audit and create the spreadsheet.

## Authors
* **Sukumar Hakhoo** - [sukumarh](https://github.com/kuramus)

## Built With
* [Visual Studio 2019](https://visualstudio.microsoft.com/vs/)
