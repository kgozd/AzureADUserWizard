
# AzureADUserWizard

This PowerShell script is designed to automate the process of creating a new user in Microsoft 365 (formerly known as Office 365) and configure various attributes for that user. The script consists of functions to connect to Microsoft 365 services, gather information about the new user and an existing user, create a new user, and display relevant information.


## Requirements

- You need to have the AzureAD and ExchangeOnlineManagement modules installed. You can install these modules using the following commands

```
Install-Module -Name AzureAD
Install-Module -Name ExchangeOnlineManagement

```

- Remember to use Windows Powershell, NOT Powershell Core

- Ensure you have the necessary permissions and credentials to connect to Microsoft 365 services.




## Usage/Examples

To use this script, execute the main function. It will guide you through the process of creating a new user in Microsoft 365, based on information you provide and an existing user's email address. After completing the process, it will display relevant user information.

Make sure to customize this script to your organization's needs and security policies. Additionally, ensure that you have the necessary permissions and access to Microsoft 365 services.

## License

[MIT](https://choosealicense.com/licenses/mit/)

