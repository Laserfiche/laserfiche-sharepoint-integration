# Test Plan for Laserfiche SharePoint Integration

## Objective
We wish to verify that changes made to the Laserfiche SharePoint Integration do not disrupt
existing functionality in the product. This plan should be executed prior to each new
release, and no changes should be included in the release until they have been tested. As 
new functionality is added to the integration, new tests should be added to the plan
to ensure adequate coverage.

## Test Cases

### Use Documentation to Install the Laserfiche SharePoint Integration sppkg
Prerequisistes:
- None

Steps:
1. Follow the instructions in the [Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation.html#deploy-laserfiche-sharepoint-integration-to-a-sharepoint-site) for adding the sppkg file to SharePoint.

Expected Result
- Instructions in documentation are clear and effective
- sppkg file ends up added to SharePoint
### Use Documentation to set up Laserfiche Sign In Page

Prerequisites:
- the SharePoint package (.sppkg) file must already be installed to SharePoint.
- follow the instructions in the README.md for running locally

Steps:
1. Follow the instructions in the [Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/adding-app-to-sp-site#) for setting up the Laserfiche Sign In Page.

Expected Result:
- Instructions in documentation are clear and effective
- the Laserfiche Sign In Page is created

### Verify Save To Laserfiche Works
Prerequisites:
- Laserfiche Sign In Page must already Exist

Steps:
1. Upload a word document with some text to a SharePoint site
1. Right click on the document
1. Select the Save To Laserfiche option
1. Select View File in Laserfiche
1. Return to the original tab and select close

Expected Result(s):
1. Each of the steps above can be done (e.g., the Save To Laserfiche Option exists for Step #3)
1. A word document is saved to Laserfiche, and its contents are identical to the original in SharePoint

### Set up Repository Explorer web part using Documentation
Prerequisites
- the SharePoint package (.sppkg) file must already be installed to SharePoint.
- follow the instructions in the README.md for running locally

Steps:
1. Follow the instructions in the [Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/adding-app-to-sp-site) for setting up the Repository Explorer Page.

Expected Result:
- Instructions in documentation are clear and effective
- the Laserfiche Repository Explorer is created

### Verify ability to navigate and open files
