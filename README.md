# Office 365 Starter Project for ASP.NET MVC #

**Table of Contents**

- [Overview](#overview)
- [Prerequisites and Configuration](#prerequisites)
- [Build](#build)
- [Project Files of Interest](#project)
- [Troubleshooting](#troubleshooting)
- [License](https://github.com/OfficeDev/Office-365-APIs-Starter-Project-for-ASPNETMVC/blob/master/LICENSE.txt)
- [Questions and Comments](#questions-and-comments)
- [Contributing](#contributing)

## Overview ##

This sample uses the [Office 365 APIs client libraries](http://aka.ms/kbwa5c) to demonstrate basic operations against the Calendar, Contacts, Mail, and Files (OneDrive for Business) service endpoints in Office 365 from a single-tenant ASP.NET MVC 5 application.  

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
  - [Microsoft Office 365 API Tools version 1.4.50428.2](http://aka.ms/k0534n). 
  - An [Office 365 developer site](http://aka.ms/ro9c62) or another Office 365 tenant.
  - Microsoft IIS enabled on your computer.

### Register app and configure the sample to consume Office 365 APIs ###

You can do this via the Office 365 API Tools for Visual Studio (which automates the registration process). Be sure to download and install the [Office 365 API tools](http://aka.ms/k0534n) from the Visual Studio Gallery before you proceed any further.

   1. Build the project. This will restore the NuGet packages for this solution. 
   2. In the Solution Explorer window, choose **O365-APIs-Start-ASPNET-MVC** project -> **Add** -> **Connected Service**.
   2. A Services Manager window will appear. Choose **Office 365** -> **Office 365 APIs** and select the **Register your app** link.
   3. If you haven't signed in before, a sign-in dialog box will appear.  Enter the user name and password for your Office 365 tenant admin. We recommend that you use your Office 365 Developer Site. Often, this user name will follow the pattern {username}@{tenant}.onmicrosoft.com. If you do not have a developer site, you can get a free Developer Site as part of your MSDN Benefits or sign up for a free trial. Be aware that the user must be a Tenant Admin user—but for tenants created as part of an Office 365 Developer Site, this is likely to be the case already. Also developer accounts are usually limited to one user.
   4. After you're signed in, you will see a list of all the services. Initially, no permissions will be selected, as the app is not registered to consume any services yet. 
   5. To register for the services used in this sample, choose the following permissions, and select the Permissions link to set the following permissions:
	- (Calendar) – Read and write to your calendars (ReadWrite)
	- (Contacts) – Read and write to your contacts (ReadWrite)
	- (Mail) - Send mail as you (Send), Read and write to your mail (ReadWrite)
	- (Files) - Read and write to your files (Write)
	- (Users and Groups) – Sign you in and read your profile (Read)
   6. Choose the **App Properties** link in the Services Manager window. Make this app available to a Single Organization. 
   7. After selecting **OK** in the Services Manager window, assemblies for connecting to Office 365 REST APIs will be added to your project and the following entries will be added to your appSettings in the web.config: ClientId, ClientSecret, AADInstance, and TenantId. You can use your tenant name for the value of the TenantId setting instead of using the tenant identifier.
   8. Build the solution. Nuget packages will be added to you project. Now you are ready to run the solution and sign in with your organizational account to Office 365.

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

Check that you are running the same version of the Microsoft Office 365 API Tools as the version used for this sample. 

Version 1.0.34 of Microsoft.Office365.OutlookServices.Portable contains a bug. This is the version installed by Microsoft Office 365 API Tools version 1.4.50428.2. Use 1.0.22 until a newer version is released.  


## Questions and Comments

We'd love to get your feedback on the O365-ASPNETMVC-Start project. You can send your questions and suggestions to us in the [Issues](https://github.com/OfficeDev/O365-ASPNETMVC-Start/issues) section of this repository.

Questions about Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Make sure that your questions or comments are tagged with [Office365] and [API].

## Contributing
You will need to sign a [Contributor License Agreement](https://cla.microsoft.com) before submitting your pull request. To complete the Contributor License Agreement (CLA), you will need to submit a request via the form and then electronically sign the Contributor License Agreement when you receive the email containing the link to the document. 



## Copyright ##

Copyright (c) Microsoft. All rights reserved.


