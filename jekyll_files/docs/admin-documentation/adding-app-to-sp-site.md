---
layout: default
title: Adding App to SharePoint Site
nav_order: 1
parent: Laserfiche SharePoint Online Integration Administration Guide
---

# Adding App to SharePoint Site

### Prerequisites

- Be an owner of a SharePoint site you want to add the integration to.

### Steps

1. Go to the SharePoint site that you would like to install the integration on. If you do not have a SharePoint site already, you can find instructions [here](https://support.microsoft.com/en-gb/office/create-a-site-in-sharepoint-4d1e11bf-8ddc-499d-b889-2b48d10b1ce8){:target="_blank"} on how to create one.
1. Open the SharePoint Store and add the app named “Laserfiche SharePoint Online Integration”. For instructions on how to add an app, reference [Microsoft's Documentation](https://learn.microsoft.com/en-us/sharepoint/use-app-catalog?redirectSourcePath=%252farticle%252fuse-the-app-catalog-to-make-custom-business-apps-available-for-your-sharepoint-online-environment-0b6ab336-8b83-423f-a06b-bcc52861cba0#add-custom-apps){:target="_blank"}.
   <a href="../assets/images/addTheApp.png"><img src="../assets/images/addTheApp.png"></a>
1. Navigate to your SharePoint site. If successfully installed, the app will be listed under the “Site contents” tab.
   <a href="../assets/images/appInstalled.png"><img src="../assets/images/appInstalled.png"></a>
**Note:** You can also [sideload the app](https://learn.microsoft.com/en-us/sharepoint/use-app-catalog?redirectSourcePath=%252farticle%252fuse-the-app-catalog-to-make-custom-business-apps-available-for-your-sharepoint-online-environment-0b6ab336-8b83-423f-a06b-bcc52861cba0#add-custom-apps){:target="_blank"} by uploading  [this solution file](../assets/LaserficheSharePointOnlineIntegration.sppkg).

### The Laserfiche Sign In Page

1. In your SharePoint site, select the "Pages" item in the navigation bar on the left side of the page.
   <a href="../assets/images/newSitePage.png"><img src="../assets/images/newSitePage.png"></a>
1. Create and open a new site page by clicking the "+ New" button and selecting "Site Page" from the dropdown.
1. Title the page “LaserficheSignIn”.
1. Move your cursor just below the title area to the white space beneath. This should reveal a hidden "+" button. If you hover over it, it should display the message "Add a new web part in column one”.
   <a href="../assets/images/hiddenPlusButton.png"><img src="../assets/images/hiddenPlusButton.png"></a>
1. Click on that button and Search for “Laserfiche Sign In".
   <a href="../assets/images/searchRepositoryExplorer.png"><img src="../assets/images/searchRepositoryExplorer.png"></a>
1. Click on the search result with a white L on an orange square. The Laserfiche Sign In web part should now appear on your Page. Before creating subsequent pages, make sure to click the 'Publish' button to save the page.

### The Repository Explorer Page

Follow the same steps as above, but title the Page whatever you wish and add the “Repository Explorer” web part.

### The Admin Configuration Page

Follow the same steps as above, but title the page whatever you wish and add the “Admin Configuration” web part instead of “Laserfiche Sign In".

### Next Steps

Before you can log in and use the web pages you just created, you will need to [Register them in the Laserfiche Developer Console](../admin-documentation/register-app-in-laserfiche). After you register your Apps, you should be able to log in and use the web parts. For Documentation on how to use the integration, reference the [User Documentation](../user-documentation/).
