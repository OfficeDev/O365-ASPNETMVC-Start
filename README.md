# Office 365 Starter Project for ASP.NET MVC #

**Table of Contents**

- [Overview](#overview)
- [Prerequisites and Configuration](#prerequisites)
- [Build](#build)
- [Project Files of Interest](#project)
- [Troubleshooting](#troubleshooting)
- [License](https://github.com/OfficeDev/Office-365-APIs-Starter-Project-for-ASPNETMVC/blob/master/LICENSE.txt)

## Overview ##

This sample uses the [Office 365 APIs client libraries](http://msdn.microsoft.com/en-us/office/office365/howto/platform-development-overview) to demonstrate basic operations against the Calendar, Contacts, Mail, and Files (OneDrive for Business) service endpoints in Office 365 from a single-tenant ASP.NET MVC 5 application.  

Below are the operations that you can perform with this sample:

**Calendar**
  - Read events
  - Add events
  - Refresh the calendar
  - Update events
  - Delete events

**Contacts**
  - Add contacts
  - Refresh the contacts list
  - Update contacts
  - Delete contacts
  
**Mail**
  - Read email messages
  - Create and send a new email

**Files (OneDrive for Business)**
  - Read files and folders.
  - Create text file.
  - Delete files and folders.
  - Read text file contents.
  - Update text file contents.
  
**Users and Groups**
  - Sign in/out

<a name="prerequisites"></a>
## Prerequisites and Configuration ##

This sample requires the following:

  - Visual Studio 2013 with Update 3.
  - [Microsoft Office 365 API Tools version 1.3.41104.1](http://aka.ms/k0534n). 
  - An [Office 365 developer site](http://aka.ms/ro9c62).
  - A subscription to [Microsoft Azure](http://aka.ms/pp13iv)
  - Microsoft IIS enabled on your computer.

### Configure the sample ###

Follow these steps to configure the sample.

   1. Open the O365-APIs-Start-ASPNET-MVC.sln file using Visual Studio 2013.
   2. Register and configure the app to consume Office 365 services (detailed below).
   3. Get your Office 365 tenant ID from Microsoft Azure (detailed below).

   Note: It is important to ensure the Office 365 API Tools are updated to the most recent version. Failure to do so may cause issues with running the sample. Again the most recent tools are located here: [Microsoft Office 365 API Tools version 1.3.41104.1](http://aka.ms/k0534n). 
### Register app to consume Office 365 APIs ###

You can do this via the Office 365 API Tools for Visual Studio (which automates the registration process). Be sure to download and install the [Office 365 API tools](http://aka.ms/k0534n) from the Visual Studio Gallery.

   1. Build the project. This will restore the NuGet packages for this solution. 
   2. In the Solution Explorer window, choose **O365-APIs-Start-ASPNET-MVC** project -> **Add** -> **Connected Service**.
   2. A Services Manager window will appear. Choose **Office 365** -> **Office 365 APIs** and select the **Register your app** link.
   3. If you haven't signed in before, a sign-in dialog box will appear.  Enter the user name and password for your Office 365 tenant admin. We recommend that you use your Office 365 Developer Site. Often, this user name will follow the pattern @.onmicrosoft.com. If you do not have a developer site, you can get a free Developer Site as part of your MSDN Benefits or sign up for a free trial. Be aware that the user must be a Tenant Admin user—but for tenants created as part of an Office 365 Developer Site, this is likely to be the case already. Also developer accounts are usually limited to one sign-in.
   4. After you're signed in, you will see a list of all the services. Initially, no permissions will be selected, as the app is not registered to consume any services yet. 
   5. To register for the services used in this sample, choose the following permissions, and select the Permissions link to set the following permissions:
	- (Calendar) – Have full access to users’ calendar
	- (Contacts) – Have full access to users’ contacts
	- (Mail) - Send mail as a user and Read and write access to users' mail
	- (Files) - Edit or delete users' files.
	- (Users and Groups) – Enable sign-on and read users’ profiles
   6. Choose the **App Properties** link in the Services Manager window. Make this app available to a Single Organization. 
   7. After selecting **OK** in the Services Manager window, assemblies for connecting to Office 365 REST APIs will be added to your project.
   8. Build the solution.

### Get your Office 365 tenant ID from Microsoft Azure ###

 In order to complete this procedure, you're going to need to log into the Microsoft Azure management portal. To do this you must have an Azure subscription. A free trial is available if you do not currently have one.
 You can sign up here: http://azure.microsoft.com/en-us/pricing/free-trial/. You must also ensure you have already completed the Register app to consume Office 365 APIs procedure.

 **Important**: You will also need to ensure your Azure subscription is bound to your Office 365 tenant. To do this see the Active Directory team's blog post, [Creating and Managing Multiple Windows Azure Active Directories](http://aka.ms/lrb3ln). The section **Adding a new directory** will explain how to do this. You can also read [Set up Azure Active Directory access for your Developer Site](http://aka.ms/fv273q) for more information.

 To retrieve your Office 365 tenant ID:

  1. Sign into the Azure management portal at https://manage.windowsazure.com/.
  2. Select the Active Directory tab in the left pane and choose your target Office 365 domain underneath the back button. As a reminder you must have the Azure subscription configured to use your specific Office 365 tenant.

	![](http://i.imgur.com/SU8Ri5f.png)

  3. Choose the Applications tab for your domain and select the registration entry for your app. It should appear as something like O365-APIs-Start-ASPNET-MVC.OfficeO365App.

	![](http://i.imgur.com/5dtWcua.png)

  4. Upon clicking that entry, expand the Enable Users To Sign On section, copy and paste the Federation Metadata Document URL value to notepad or another application. You'll notice that there's an identifier present in that URL (in the form of a guid), and this is the tenant ID that is needed for the project. 

	![](http://i.imgur.com/TzXIlut.png)

  5. Copy just the identifier value and return to the sample solution.
  6. Add your tenant ID to the ida:TenantID key in the web.config. It should look similar to this:
	 `<add key="ida:TenantID" value="d10f81ac-2de0-4eaf-af91-393d1bdaf17d"/>`
  7. You are now ready to build the project.
  
Note: If you are deploying to a production tenant, you will need to ask your tenant admin for the tenant identifier.
  
## Build ##

After you've loaded the solution in Visual Studio, press F5 to build and debug.
Run the solution and sign in with your organizational account to Office 365.

<a name="project"></a>
## Project Files of Interest ##

**Controllers**
   - AccountController.cs
   - CalendarController.cs
   - ContactController.cs
   - FileController.cs
   - HomeController.cs
   - MailContoller.cs

**Helper Classes**
   - AuthenticationHelper.cs
   - CalendarOperations.cs
   - ContactOperations.cs
   - FileOperations.cs
   - MailOperations.cs
 
**Models**
   - CalendarEvent.cs
   - ContactItem.cs
   - FileItem.cs
   - IdentityModels.cs
   - MailItem.cs

**Utils Folder** 
   - SettingsHelper.cs
   - NaiveSessionCache.cs - This is a sample token cache and should not be used in a production environment. We suggest that you store and interact with tokens in accordance with the security policy of your organization. 

**Views**
   - Calendar/Create.cshtml
   - Calendar/Delete.cshtml
   - Calendar/Edit.cshtml
   - Calendar/Index.cshtml
   - Contact/Create.cshtml
   - Contact/Delete.cshtml
   - Contact/Edit.cshtml
   - Contact/Index.cshtml
   - File/Create.cshtml
   - File/Delete.cshtml
   - File/Edit.cshtml
   - File/Index.cshtml
   - Home/Index.cshtml
   - Mail/Create.cshtml
   - Mail/Delete.cshtml
   - Mail/Index.cshtml
   - Shared/_Layout.cshtml
   - Shared/_LoginPartial.cshtml

**Other**
   - RouteConfig.cs
   - web.config
   - Startup.cs
   - Startup.Auth.cs
   - packages.config

## Troubleshooting ##

If you see any errors while installing packages, for example, *Unable to find "Microsoft.Azure.ActiveDirectory.GraphClient" version="1.0.21"*, make sure the local path where you placed the solution is not too long/deep. Moving the solution closer to the root of your drive resolves this issue. We'll also work on shortening the folder names in a future update. There is a long file name restriction of about 260 characters in Visual Studio. 

Your browser will not display a web page if you try to sign-in and the application doesn't have the Users and Groups,  **Enable sign-on and read users’ profiles** option selected. 

The **Specified argument was out of the range of valid values. Parameter name: site** will occur if IIS is not enabled. 

An incorrect tenant identifier will return a 404 HTTP status code. 
## Copyright ##

Copyright (c) Microsoft. All rights reserved.


