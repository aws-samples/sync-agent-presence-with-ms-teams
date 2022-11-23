# Sync Amazon Connect agent availability status with Office 365 Microsoft Teams
This demo shows how-to enable the syncing of agent statuses from Amazon Connect to Office 365 Microsoft Teams.

## 🎒 Pre-requisites

Follow these steps to create the Office 365 application. If you don’t have an existing office 365 account for testing, you can use the [free trial of Office 365 business premium](https://www.microsoft.com/en-us/p/office-365-business-premium/CFQ7TTC0K5J7).

**Note:** 
To complete setup, you might need to work with your Azure Active Directory administrator.

## Configure Office 365

Log in to the Azure portal [https://portal.azure.com/](https://portal.azure.com/) with your Office365 account credentials. Once logged in, click on Azure Active Directory.

1. In the Azure Active Directory console, click on App registrations. In the app registration screen, click on new registration.
2. On the new registration screen, enter application name, select your choice of account type and click on Register.
3. Make note of the **Application (client) ID** and **Directory (tenant) ID**
4. Click on View API Permissions
5. In the API Permissions page, click on Add a permission
6. Select Microsoft Graph and click on Application permissions
7. In the Select permissions section, search for **Presence.ReadWrite.All, User.Read.All** and click on Add permissions
8. Click on grant admin consent. In the confirmation pop up, click Yes to grant Admin consent for this app
9. In the left navigation, Click on certificates and secrets
10. Click on New client secret and select expiration period and click on Add
11. Copy the **Client Secret Value** and store it securely

**Note:** 
Make sure you copy values from steps 3 and 11, you will need them to deploy CloudFormation

## Launch Stack
Click on the launch stack link to deploy the solution, make sure you are using the same region where you deployed your Amazon Connect instance 

[Launch Stack](https://console.aws.amazon.com/cloudformation/home#/stacks/create/review?templateURL=https://aws-contact-center-blog.s3.us-west-2.amazonaws.com/sync-agent-presence-with-ms-teams/cloudformation/o365.yaml) 

On the AWS CloudFormation console, enter **Application (client) ID** , **Directory (tenant) ID** and **Client Secret Value**
CloudFormation will securely store these values in your account using AWS Secrets Manager.
**Note:** 
Optionally, the CloudFormation is available for download from the [cloudformation folder](./cloudformation)

## Configure Amazon Connect

1. [Add o365 Lambda function](https://docs.aws.amazon.com/connect/latest/adminguide/connect-lambda-functions.html) that was created by CloudFormation to your Amazon Connect instance
2. Import [contact flows](https://docs.aws.amazon.com/connect/latest/adminguide/contact-flow-import-export.html) from [contact flows folder](./contact-flows) starting with o365_agent_connect, o365_agent_disconnect and then import o365_main
3. [Assign the contact flow](https://docs.aws.amazon.com/connect/latest/adminguide/tutorial1-assign-contact-flow-to-number.html) name *o365_main* to the phone number
4. Login as an agent, place and answer a call, review your office 365 status in MS Teams,  disconnect,  close the contact in Amazon Connect CCP, and review your office 365 status in MS Teams again

**Note:** 
Lambda function is using agent email to [look-up user id in office 365](https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http), this could be customized based on your needs by updating or changing **lambda_function.py** ***o365_get_user_by_email*** **function**

## 🏛️ License
This library is licensed under the MIT-0 License. See the LICENSE file.
