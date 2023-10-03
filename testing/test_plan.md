# Test Plan for Laserfiche SharePoint Online Integration

## Objective

Verify that changes made to the Laserfiche SharePoint Online Integration do not disrupt
existing functionality in the product. This plan should be executed prior to each new
release, and no changes should be included in the release until they have been tested. As
new functionality is added to the integration, new tests should be added to the plan
to ensure adequate coverage.

## Test Cases

- [Installation](#installation)
- [Site Configuration](#site-configuration)
- [Integration Configuration](#integration-configuration)
- [Save to Laserfiche](#save-to-laserfiche)
- [Repository View](#repository-view)

### **Installation**

#### **Use Documentation to Add the Integration to your Tenant App Catalog**

Steps:

1. Follow the instructions in the [Adding App to Organization Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/adding-app-organization.html)

Expected Results:

- Instructions in documentation are clear and effective
- `Laserfiche SharePoint Online Integration` is available in your tenant app catalog

#### **Use Documentation to Add the Integration to your SharePoint Site**

Steps:

1. Follow steps 1-5 in the [Adding App to SharePoint Site Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/adding-app-to-sp-site.html)

Expected Results:

- Instructions in documentation are clear and effective
- `Laserfiche SharePoint Online Integration` is available in your SharePoint Site site contents

### Site Configuration

Prerequisites:

- Follow the [Installation](#installation) steps successfully

#### Use Documentation to set up Laserfiche Sign In Page

Steps:

1. Follow the instructions for Laserfiche Sign in page in the [Add Add to SP Site Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/add-app-to-sp-site)

Expected Results:

- Instructions in documentation are clear and effective
- You have a page in your site called LaserficheSignIn that contains the Laserfiche Sign In Web Part

#### Use Documentation to set up Laserfiche Repository Explorer

Steps:

1. Follow the instructions for the Repository Explorer page in the [Add Add to SP Site Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/add-app-to-sp-site)

Expected Results:

- Instructions in documentation are clear and effective
- You have a page in your site that contains the Laserfiche Repository Explorer Web Part

#### Use Documentation to set up Laserfiche Admin Configuration

Steps:

1. Follow the instructions for Admin Configuration page in the [Add Add to SP Site Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/add-app-to-sp-site)

Expected Results:

- Instructions in documentation are clear and effective
- You have a page in your site that contains the Laserfiche Administrator Configuration Web Part

#### Use Documentation to register app in dev console

Steps:

1. Follow the instructions for registering application in the [Register App in Laserfiche Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/register-app-in-laserfiche.html)

Expected Results:

- Verify manifest is valid (i.e. SPA, correct clientId, etc.)
- App is registered successfully in Laserfiche dev console
- You can sign in on all three pages you created above

### Integration Configuration

Prerequisites:

- Follow the [Installation](#installation) and [Site Configuration](#site-configuration) steps successfully

#### Create standard profile

Prerequisites:

- the admin configuration web part must exist in a SharePoint Page
- finish testing the functionality of the repository explorer web part

Steps:

1. Go to the Profiles tab and click the `Add Profile` button.
1. Name the Profile `Example Profile Name`, do not select a template, select the Folder which you created in the functionality test of the Repository Explorer web part as the destination folder, and choose `Leave a copy of the file in SharePoint` for the `After import` behavior. Click the Save button.
1. Go to the Profile Mapping tab and click the `Add` button.
1. Select `Document` for the SharePoint Content Type and select `Example Profile Name` for the `Laserfiche Profile`. Click the floppy disk icon to save.

Expected Results:

1. Something resembling the following Profile Editor appears: [Could Not Display Image](./assets/profileCreator.png)
1. You should get a Success dialog, and then get returned to the `Profiles Tab`, where the new profile should be visible.

### Save to Laserfiche

Prerequisites:

- Laserfiche Sign In Page must already Exist

#### Test happy path save

Steps:

1. Upload a document of some kind with some text to the Document's tab of a SharePoint site
1. Right-click on the document.
1. Select the Save To Laserfiche option
1. Select View File in Laserfiche
1. Return to the original tab and select close

Expected Results:

1. Does not test Integration behavior
1. The `Save to Laserfiche` option should exist in the resulting drop down.
1. A dialog should immediately open, and eventually display a success message and a button saying `View File in Laserfiche`
1. The file should be opened in a new tab with a `Back` button. Clicking the `Back` button should display a folder view containing the document saved to Laserfiche.
1. The dialog should disappear.

### Repository View

Prerequisites

- Repository Explorer web part must exist

Steps:
1. Follow the instructions in the [Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/adding-app-to-sp-site) for setting up the Repository Explorer Page.

Expected Results:

1. We should be able to see the new folder without needing to reload anything
1. We should see an empty set of children and the breadcrumb navigation should appear
1. We should be able to see the new folder without needing to reload anything. Upon double-clicking, we should see an empty set of children and the breadcrumb navigation should still allow us to move back up to the ancestor folders.
1. A dialog should pop up allowing you to choose the file. After you've chosen, the file should be immediately visible.
1. the selected file row should be visibly distinguished from the other non-selected rows. Additionally, clicking on the arrow-in-square should open a new tab to view the file in Laserfiche.
1. the breadcrumb navigation links should take us to the named locations.
