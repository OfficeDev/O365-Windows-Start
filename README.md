# Office 365 Starter Project for Windows Store App #

**Table of Contents**

- [Overview](#overview)
- [Prerequisites and Configuration](#prerequisites)
- [Build](#build)
- [Project Files of Interest](#project)
- [Troubleshooting](#troubleshooting)
- [License](https://github.com/OfficeDev/Office-365-APIs-Starter-Project-for-Windows/blob/master/LICENSE.txt)

## Overview ##

The Office 365 Starter Project sample uses the Office 365 API Tools client libraries to illustrate basic operations against the Calendar and Contacts service endpoints in Office 365.
The sample also demonstrates how to authenticate against multiple Office 365 services in a single app experience, and retrieve user information from the Users and Groups service.
We'll be adding examples of how to use Files and Email client APIs in updates to this project, so make sure to check back.
Below are the operations that you can perform with this sample:

Calendar  
  - Get calendar events  
  - Create events  
  - Update events  
  - Delete events  

Contacts  
  - Get contacts  
  - Create contacts  
  - Update contacts  
  - Delete contacts  
  - Change contact photo  

Users and Groups  
  - Get display name  
  - Get job title  
  - Get profile picture  
  - Get user ID  
  - Check signed in/out state  

<a name="prerequisites"></a>
## Prerequisites and Configuration ##

This sample requires the following:
  - Visual Studio 2013 with Update 3.
  - [Office 365 API Tools version 1.2.41027.2](https://visualstudiogallery.msdn.microsoft.com/7e947621-ef93-4de7-93d3-d796c43ba34f).
  - An [Office 365 developer site](https://portal.office.com/Signup/Signup.aspx?OfferId=6881A1CB-F4EB-4db3-9F18-388898DAF510&DL=DEVELOPERPACK&ali=1).
  - At least one sign-on to OneDrive for Business via the web browser.

###Configure the sample

Follow these steps to configure the sample.

   1. Open the O365-APIs-Start-Windows.sln file using Visual Studio 2013.
   2. Register and configure the app to consume Office 365 services (detailed below).

###Register app to consume Office 365 APIs

You can do this via the Office 365 API Tools for Visual Studio (which automates the registration process). Be sure to download and install the Office 365 API tools from the Visual Studio Gallery.

   1. In the Solution Explorer window, choose Office365Starter project -> Add -> Connected Service.
   2. A Services Manager dialog box will appear. Choose Office 365 and Register your app.
   3. On the sign-in dialog box, enter the username and password for your Office 365 tenant. We recommend that you use your Office 365 Developer Site. Often, this user name will follow the pattern <your-name>@<tenant-name>.onmicrosoft.com. If you do not have a developer site, you can get a free Developer Site as part of your MSDN Benefits or sign up for a free trial. Be aware that the user must be an Tenant Admin user—but for tenants created as part of an Office 365 Developer Site, this is likely to be the case already. Also developer accounts are usually limited to one sign-in.
   4. After you're signed in, you will see a list of all the services. Initially, no permissions will be selected, as the app is not registered to consume any services yet. 
   5. To register for the services used in this sample, choose the following permissions:  
   	- (Calendar) – Have full access to users’ calendar  
	- (Contacts) - Have full access to users' contacts  
	- (Users and Groups) – Enable sign-on and read users' profiles  
   6. After clicking OK in the Services Manager dialog box, assemblies for connecting to Office 365 REST APIs will be added to your project.

## Build ##

After you've loaded the solution in Visual Studio, press F5 to build and debug.
Run the solution and sign in with your organizational account to Office 365.

<a name="project"></a>
## Project Files of Interest ##

**Helper Classes**  
   - CalendarOperations.cs  
   - UserOperations.cs  
   - AuthenticationHelper.cs  
   - ContactOperations.cs  

**View Models**  
   - CalendarViewModel.cs  
   - EventViewModel.cs  
   - UserViewModel.cs  
   - ContactsViewModel.cs  
   - ContactItemViewModel.cs  

## Troubleshooting ##

You may run into an authentication error after deploying and running if apps do not have the ability to access account information in the [Windows Privacy Settings](http://www.microsoft.com/security/online-privacy/windows.aspx) menu. Set **Let my apps access my name, picture, and other account info** to **On**. This setting can be reset by a Windows Update. 

## Copyright ##

Copyright (c) Microsoft. All rights reserved.

