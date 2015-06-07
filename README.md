Office 365 Contacts Single Page Application using VisualStudio
This sample shows how to build a AngularJS Single Page Application using Visual Studio for office 365 Contacts API 

ADAL for Javascript is an open source library. For distribution options, source code, and contributions, check out the ADAL JS repo at https://github.com/AzureAD/azure-activedirectory-library-for-js.

How To Run This Sample

To run this sample, you need:

Visual Studio 2013
Office Developer Tools for Visual Studio 2013
Office 365 Developer Subscription
Step 1: Clone the application in Visual Studio

Visual Studio 2013 supports connecting to Git servers. As the project templates are hosted in GitHub, Visual Studio 2013 makes it easier to clone projects from GitHub.

The steps below will describe how to clone Office 365 API web application project in Visual Studio from Office Developer GitHub.

Open Visual Studio 2013.
Switch to Team Explorer.
Team Explorer provides options to clone Git repositories.
Click Clone under Local Git Repositories, enter the clone URL https://github.com/girishgoudar/Office365SPAContacts.git for the web application project and click Clone.
Once the project is cloned, double click on the repo.
Double click the project solution which is available under Solutions.
Switch to Solution Explorer.
Step 2: Configure the sample

Build the Project

Simply Build the project to restore NuGet packages.

Register Azure AD application to consume Office 365 APIs

Office 365 applications use Azure Active Directory (Azure AD) to authenticate and authorize users and applications respectively. All users, application registrations, permissions are stored in Azure AD.

Using the Office 365 API Tool for Visual Studio you can configure your web application to consume Office 365 APIs.

In the Solution Explorer window, right click your project -> Add -> Connected Service.
A Services Manager dialog box will appear. Choose Office 365 -> Office 365 API and click Register your app.
On the sign-in dialog box, enter the username and password for your Office 365 tenant.
After you're signed in, you will see a list of all the services.
Initially, no permissions will be selected, as the app is not registered to consume any services yet.
Select Users and Groups and then click Permissions
In the Users and Groups Permissions dialog, select Enable sign-on and read users profiles' and click Apply
Select My Files and then click Permissions
In the My Files Permissions dialog, select both Read users' files and Edit or delete users' files then click Apply
Click Ok
After clicking OK, Office 365 client libraries (in the form of NuGet packages) for connecting to Office 365 APIs will be added to your project.

In this process, Office 365 API tool registered an Azure AD Application in the Office 365 tenant that you signed in the wizard and added the Azure AD application details to web.config.

Update web.config with your Tenant ID

In your web.config, update the TenantId value to your Office 365 tenant Id where the application is deployed.

To get the tenant Id of your Office 365 tenant:

Log in to your Azure Portal and select your Office 365 domain directory.
Note: If you are unable to login to Azure Portal using your Office 365 credentials, You can also access your Office 365’s Azure Portal directly from your Office 365 Admin Center

Now, in the browser URL, locate the GUID. This will be your Office 365 tenant Id.
Copy and paste it in the web.config where it says “paste-your-tenant-guid-here“ :
Note: If you are deploying to a production tenant, you will need to ask your tenant admin for the tenant identifier.

Update app.js with your Client ID

Open the web.config file.
Find the app key ida:ClientID and copy the value.
Open the file App/Scripts/app.js and locate the line adalProvider.init(.
Replace the value of clientId with the ClientId from web.config.
Step 3: Enable the OAuth2 implicit grant for your application

By default, applications provisioned in Azure AD are not enabled to use the OAuth2 implicit grant. In order to run this sample, you need to explicitly opt in.

Log in to your Azure Portal and select your Office 365 domain directory.
Click on the Applications tab.
Paste the ClientID which was copied from web.config into the search box and trigger search. Now select "Office365CORS.Office365App" application from the results list.
Using the Manage Manifest button in the drawer, download the manifest file for the application and save it to disk.
Open the manifest file with a text editor. Search for the oauth2AllowImplicitFlow property. You will find that it is set to false; change it to true and save the file.
Using the Manage Manifest button, upload the updated manifest file. Save the configuration of the app.
Step 4: Run the sample

Clean the solution, rebuild the solution, and run it.

In the Debug Toolbar, select to run with Google Chrome instead of Internet Explorer.

Note: This sample only works on current versions of Chrome and Firefox. A separate patch is required to be installed if you want to try this out in Internet Explorer. Links to the patch will be added soon.

You can trigger the sign in experience by either clicking on the Login link on the top right corner, or by clicking directly on the My Files tab. To see a list of files stored on OneDrive click on the My Files tab. To view the properties of a file or folder, click on the properties link displayed right next to the file name and the properties will be displayed at the bottom of MyFiles list.

Step 5: Deploy This Sample To Azure Website

To deploy this sample application to azure website, we will first create a new website on Azure then import the publish profile

Right Click on Office365SPAContacts project node in Solution Explorer.
Select Publish.
On the Publish Web wizard, select Microsoft Azure Websites.
Sign into your azure subscription if your are not currently signed in.
On the "Select Existing Website" dialog, click New....
Enter a site name, for example "Office365CORS".
Select a Region.
Let Database server be set to "No database".
Click Create.
In Connection tab, make sure that Publish method is "Web Deploy". Click Next.
In Settings tab, un-check the Enable Organizational Authentication checkbox to disable it.
NOTE: Check this option if you are deploying your Azure AD application to a different domain than the one used to build your web application. In this sample, we are deploying to the same domain for testing.
Click Publish.
Sign in to the Azure management portal.
Select the newly created website and click on Configure tab.
Scroll down to Authentication / Authorization and click Configure.
On the "Configure Azure Wedsites Authentication / Authorization" dialog, select the Directory in which the sample app was registered and from the AAD Application drop down list, select "Office365CORS.Office365App (https://localhost:44304/)". Click OK.
The sample application will now be published to "[websiteURL].azurewebsites.net".
