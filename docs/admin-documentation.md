---
layout: default
title: Admin Documentation
nav_order: 2
---

## Deploy Laserfiche Apps to SharePoint Sites

1. [Click here](https://go.microsoft.com/fwlink/?linkid=2185219) to go to the SharePoint Admin Center or find the same link at [learn.microsoft.com](https://learn.microsoft.com/en-us/sharepoint/sharepoint-admin-role#about-the-sharepoint-administrator-role-in-microsoft-365). If you can't access this webpage, then you need to ask an Administrator for help.
1. In the navigation menu on the left, select the "More features" item.
1. Open "Apps".
1. Click Upload and select the solution file. If you don't have the solution file, then clone the [SharePoint Repo](https://github.com/Laserfiche/laserfiche-sharepoint-integration) and follow the instructions in the README.md for building the
solution.
1. In your SharePoint Site (Not the Admin Center), navigate to your
site’s App catalog by clicking on the "Site Contents" item in the
navigation bar on the left side of the page.
1. Open the "New" Dropdown menu by clicking on the "+" icon.
1. Add the App named “laserfiche-sharepoint-integration-client-side-solution” if it is present in the list of Apps you can add, or look for it in the SharePoint Store. 
1. Enable the app if you are asked to do so.
1. Return to your site. If you can see  the SharePoint integration app under the “Site Contents” tab, then you have successfully installed the app.


## Use Laserfiche Apps on SharePoint Pages

- The Repository Access Web Part:
    1. In your SharePoint Site, select the "Pages" item in the navigation bar on the left side of the page.
    1. Create and open a new site page by clicking the blue "+ New" button and selecting "Site Page" from the dropdown.
    1. Title the page “LaserficheSpApp”.
    1. Move your cursor just below the title area to the white space beneath. This should reveal a hidden "+" button. If you hover over it, it should display the message "Add a new web part in column one”.
    1. Click on that button and Search for “Repository Access”.
    1. Click on the search result with a white L on an orange square. The Repository Access WebPart should now appear on your Page. Before using the Webpart, make sure to [Register Your App in the Laserfiche Developer Console](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation.html) that link aint to good right now
    1. After you register your App and give it the right redirect URI, you should be able to log in and use the component. For Documentation on how to use the components, reference the [User Documentation](./user-documentation.html)

- The Admin Configuration Web Part:
    - Follow the same steps as above, but title the Page“LaserficheSpAdministration”, and add the “Admin Configuration”web part instead of “Repository Access".
- The Save-To-Laserfiche Web Part:
    - Follow the same steps as above, but title the Page LaserficheSpSignIn, and add the “Save to Laserfiche” web part.

## How to Register Your App in the Developer Console
1. Open the [Developer Console](https://developer.laserfiche.com/developer-console.html)
1. Attempt to Create a New App from Manifest, and copy-paste the manifest provided [here](https://github.com/Laserfiche/laserfiche-sharepoint-integration/blob/1.x/UserDocuments/Laserfiche%20SharePoint%20Integration%20AppManifest.json)
1. If the attempt fails because an app with that client ID already exists, find the app with that client id by opening the following url in a new tab: https://app.laserfiche.com/devconsole/apps/<b>{your_client_id_goes_here}</b>/config, where the part in brackets should be replaced by the client_id you copied.
1. One way or another, an app with that client ID should now exist. Open the app in devconsole and switch from the general tab to the authentication tab.
1. Add the url of your SharePoint Page with the Laserfiche Web Part as a new redirect URI.
1. You should now be able to Sign In.