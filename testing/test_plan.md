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
- [Admin Access Rights](#admin-access-rights)
- [Profiles](#profiles)
- [Save to Laserfiche and Profile Mapping](#save-to-laserfiche-and-profile-mapping-tab)
- [Repository Explorer](#repository-explorer)

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

### Admin Access Rights

Prerequisites:

- Follow the [Installation](#installation) and [Site Configuration](#site-configuration) steps successfully
- You must NOT BE a site owner of the site containing that page.

#### Test Restricted Access to Web Part

Steps:
1. Attempt to open the admin configuration web part on the protected page.
    - Expected Results: You should not be able to do any configuring, and there should be an error message explaining that you don't have the necessary rights.


### Profiles
Prerequisites:

- Follow the [Installation](#installation) and [Site Configuration](#site-configuration) steps successfully
- You must BE a site owner of the site containing that page.
#### Create standard profile
Steps:
1. Go to the Profiles tab and click the `Add Profile` button.
1. Name the Profile `Example Profile Name`, do not select a template, select a folder you have access to save into, and choose `Leave a copy of the file in SharePoint` for the `After import` behavior. Click the Save button.
   - Expected Results: You should see a success dialog, and then get returned to the `Profiles Tab`, where the new profile should be visible.


#### Test Profile Error Handling
Steps:
1. Go to the Profiles tab and click the `Add Profile` button.
1. Name the Profile, `Bad Profile`, select the Folder which you created in the functionality test of the Repository Explorer web part as the destination folder, and select the `General` template. In the Mapping section, Click `Add Field`, and choose `Actual Work` for the SharePoint Column and `Date (2)` for the Laserfiche Field.
    - Expected Results: 
      - Warning/error that the data types don't match.
      - Profile is not saved.
1. Delete the SP Column/LF field pair.
    - Expected Results:
      - Profile can be saved (button not disabled)
1. Save the Profile
1. Add a New Profile, and name it `Bad Profile` as well. Attempt to Save.
    - Expected Results:
        - The Profile should not be added.
        - The page should not indicate that the profile was added
        - The page should explain that the profile was not added because a profile with that name already exists.
1. Delete the profile named `Bad Profile`.
#### Test Edit Profile

Steps:
1. Click the pencil button to edit a profile and add some compatible metadata mappings like a text type for the SharePoint Column and a String type for the Laserfiche Field, for example. Click Save.
    - Expected Results:
      - The page should indicate that the profile was saved.

#### Test 'after import' configuration

Steps:
1. Create a Profile named `Duplicate in Laserfiche` that saves to a folder of your choice and leaves a copy of the file in SharePoint after import.
1. Create a Profile named `Replace with Link` that saves to the same folder and Replaces SharePoint file with a link after import.
1. Create a Profile named `Delete From SharePoint` that saves to the same folder and Deletes SharePoint file after import.

Expected Results:
  - Those three profiles exist

#### Test metadata configuration
Steps:
1. Create a Profile named `number metadata` that saves to a folde of your choice and leaves a copy of the file in SharePoint after import. Assign a template that has a required number field in Laserfiche, and map the SharePoint Column `Actual Work` to the required number field.
1. Save the Profile

Expected Results:
- Profile appears


### Save to Laserfiche and Profile Mapping Tab

Prerequisites:

- Follow the [Profiles](#profiles) Tests successfully
- Laserfiche Sign In Page must already exist
- Laserfiche Admin Configuration Page must already exist
#### Test Default Profile with No Content Type 
Steps:
1. Inside the SharePoint site's `Documents` tab, remove the column displaying `Content Type`. 
1. In the Profile Mapping Tab, associate the `[Default]` SharePoint Content Type with the `Example Profile Name` Laserfiche Profile. Remember to Save the mapping. Make sure no other mappings exist.
1. Attempt to save a file to Laserfiche inside the Documents tab.

Expected Results
- The file is saved according to the Default Profile
#### Test Default Profile
Steps:
1. Replace content type as a column in the Documents tab.
1. In the Profile Mapping Tab, associate the `[Default]` SharePoint Content Type with the `Example Profile Name` Laserfiche Profile. Remember to Save the mapping.
1. Eliminate all other mappings
1. Save a document from the Documents Tab of the SharePoint site to Laserfiche

Expected Results:
  - The file should save in the destination folder you configured in the Default section.

#### Test No Default Profile Save
Steps:
1. Remove all SharePoint Content Type -> Laserfiche Profile Mappings
1. Attempt to save a Document from the documents tab

Expected Result:
  - Document does not save
  - Error that requests a default mapping or a mapping for the relevant content type

#### Test Save when a required field doesn't exist
Steps:
1. Add SharePoint Column `Actual Work` to SharePoint Library
1. Make sure that `Actual Work` has no value for a specific document
1. Set the Default mapping to `number metadata`, and save. There should be no other mappings
1. Attempt to save the specific document to Laserfiche

Expected Results:
  - The document does not save
  - Error message that says that Actual Work doesn't have a value.
#### Test metadata constraint failed case
Steps:
1. Add SharePoint Column "Actual Work" to SharePoint Library
1. Add value for "Actual Work" for a specific document to be a very large number
1. Set the Default mapping to `number metadata`, and save.
1. Attempt to save the specific document to Laserfiche
  
Expected Results:
  - The document should save, BUT
  - There should be a warning that says the metadata didn't save.

#### Test specific mapping overrides default
Steps:
1. Remove all Profile Mappings
1. Add a mapping from `[Default]` to `Default`
1. Add a mapping from `Document` to `number metadata`
1. Choose a Document in SharePoint
1. Update the `Actual Work` column of the document to have a value of 5
1. Attempt to save the document to Laserfiche

Expected Results:
  - The document should successfully save
  - The document's number field should have a value of 5.
#### Test replace with URL action
Steps:
1. Edit the mapping from `Document` so that it points to `Replace with Link`
1. Attempt to save a Document to Laserfiche

Expected Results:
  - The document should successfully appear in Laserfiche
  - In SharePoint, the document should be replaced with a link
  - Link should actually link to the document in LF

#### Test delete after save to Laserfiche
Steps:
1. Edit the mapping from  `Document` so that it points to `Delete From SharePoint`
1. Attempt to save a Document to Laserfiche

Expected Results:
  - The document should exist in Laserfiche and no longer exist in SharePoint

#### Test saving .url files to Laserfiche
Steps:
1. Attempt to save a .url file to Laserfiche

Expected Result:
  - You should be told that you can't save a .url file to Laserfiche.

#### Test Mapping Content Types to multiple Profiles
Steps:
1. In addition to the existing `[Default]` -> `Default` mapping, add a mapping from `[Default]` to `Replace with Link`.
1. Click Save

Expected Results
  - The new mapping should not save
  - You should see an error message saying a mapping already exists for that content type
### Repository Explorer

Prerequisites:

- Follow the [Installation](#installation) and [Site Configuration](#site-configuration) steps successfully

#### Test login

Steps:

1. Click `Sign in` button
   - Expected results: You are led through the OAuth flow, you return to repository explorer page, and button says `Sign Out`

#### Test Open button

Steps:

1. Refresh repository explorer to the root folder
1. Click open button
   - Expected result: Open root folder in Laserfiche in a new tab
1. Return to repository explorer, double-click on a folder to enter it.
1. Select (single-click) a folder inside
1. Click open button
   - Expected result: Open the selected folder in Laserfiche in a new tab
1. Select (single-click) a document inside
1. Click the open button
   - Expected result: Open the selected document in Laserfiche in a new tab

#### Test import file button

Steps:

1. Navigate to a folder that you have access to create documents in
1. Have no folder/document selected
1. Click the import file button
1. Click import without uploading file
    - Expected behavior: Error message stating please select a file to upload
1. Upload test file using browse button
1. Add no metadata
1. Click ok
    - Expected behavior: Dialog closes
1. Use refresh button to refresh open folder
    - Expected behavior: File exists in currently opened folder
1. Back in repository explorer, single-click a folder
1. Click the import file button
1. Upload test file using browse button
1. Add no metadata
1. Click ok
    - Expected behavior: Dialog closes
1. Use refresh button to refresh open folder
    - Expected behavior: File exists in currently opened folder (not the one selected)
1. Back in the repository explorer, single-click a file
1. Click the import file button
1. Upload test file using browse button
1. Add no metadata
1. Click ok
    - Expected behavior: Dialog closes
1. Use refresh button to refresh open folder
    - Expected behavior: File exists in currently opened folder (not the one selected)
1. Back in repository explorer, click the import file button
1. Upload test file using browse button
1. Add template
1. Make an error in the metadata
1. Attempt to upload file
    - Expected behavior: File not uploaded, metadata component shows relevant errors if not already shown
1. Add valid metadata
1. Click ok
    - Expected behavior: Dialog closes
1. Use refresh button to refresh open folder
    - Expected behavior: File exists in currently opened folder
1. Double-click recently imported file
    - Expected behavior: Metadata specified was successfully set
1. Back in repository explorer, click the import file button
1. Upload test file using browse button
1. Rename file to be same as existing document
1. Click ok
    - Expected behavior: Dialog closes
1. Use refresh button to refresh open folder
    - Expected behavior: File was uploaded, but has been automatically renamed

#### Test Create folder button

Test delete after save action 

1. Navigate folder where you have permissions to create entries
1. Use create folder button
1. Create folder with valid name
    - Expected Results: Dialog closes
1. Use refresh button
    - Expected results: New folder exists in currently open folder
1. Use create folder button
1. Attempt to create with no name
    - Expected results: Dialog remains open, error specifies to provide a folder name
1. Close dialog
1. Use create folder button
1. Use name with invalid characters (Ex/ )
1. Attempt to create folder
    - Expected results: Dialog remains open, error  specifies to provide a valid folder name
1. Use create folder button
1. Use name that already exists in folder
1. Attempt to create
    - Expected Results: Dialog remains open, receive error that object already exists
1. Select (single-click) a folder in the repository explorer
1. Use create folder button
1. Create folder with valid, unique name
    - Expected Results: Dialog closes
1. Use refresh button
    - Expected results: New folder exists in currently open folder

#### Test refresh button

Steps:

1. Open specific folder in repository explorer
1. Open same folder in Web Client in a new tab
1. Create folder in Web Client in that folder
1. Return to repository explorer tab
1. Click refresh button
    - Expected behavior: Folder that was created in Web Client will now exist in the repository explorer
