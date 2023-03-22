import * as React from "react";
require('../../../../Assets/CSS/adminConfig.css');
require('../../../../Assets/CSS/bootstrap.min.css');

export default class HomePage extends React.Component {
    public render(): React.ReactElement {
        return (
            <div className="container-fluid p-3">
                <main className="bg-white shadow-sm" style={{"width":"80%"}}>
                    <div className="p-3">
                        <p className="adminContent">The Laserfiche Administration page lets you edit your SharePoint and Laserfiche configuration. Sign in and select the task you want to perform from the menu on the top of this section.</p>
                        <p className="adminContent">For more information, see the <a href="https://doc.laserfiche.com/laserfiche.documentation/11/administration/en-us/Default.htm#../Subsystems/Integrations/Content/SharePoint/SharePoint2022Integration.htm" target='_blank'>help documentation.</a> <i>Note: the help link is not live yet.</i> </p>
                        <div className="row mt-3">
                            <div className="adminContent">
                                <p><strong>Profiles</strong></p>
                                <p style={{ "marginLeft": "38px" }}><span>Profiles govern how documents in SharePoint will be saved to Laserfiche. You can create multiple profiles for different SharePoint content types. For example, you may want applications stored differently than invoices, and thus you’ll create separate profiles for each.
                                </span></p>
                                <p><strong>Profile Mapping</strong></p>
                                <p style={{ "marginLeft": "38px" }}><span>In this tab, you can map a specific SharePoint content type with a corresponding Laserfiche profile. This profile will then be used when saving all documents of the specified SharePoint content type.
                                </span></p>
                            </div>
                        </div>
                    </div>
                </main>
            </div>
        );
    }
}