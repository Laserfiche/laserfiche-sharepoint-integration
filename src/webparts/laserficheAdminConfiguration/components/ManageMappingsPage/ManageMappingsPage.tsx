import * as React from "react";
import * as $ from 'jquery';
import { IManageMappingsPageProps } from "./IManageMappingsPageProps";
import { IManageMappingsPageState } from "./IManageMappingsPageState";
import { IListItem } from './IListItem';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
require('../../../../Assets/CSS/bootstrap.min.css');
require('../../../../Assets/CSS/adminConfig.css');
require('../../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');

export default class ManageMappingsPage extends React.Component<IManageMappingsPageProps, IManageMappingsPageState>{
    constructor(props: IManageMappingsPageProps) {
        super(props);
        this.Reset=this.Reset.bind(this);
        this.state = {
            mappingRows: [],
            sharePointContentTypes: [],
            laserficheContentTypes: [],
            listItem: [],
            showDeleteModal: false,
            deleteSharePointcontentType: ''
        };
    }
    //On component load get content types from SharePoint and get existing mapping list from the SharePoint Admin Configuration list
    public componentDidMount(): void {
        this.setState(() => { return { showDeleteModal: false }; });
        this.GetAllSharePointContentTypes();
        this.GetAllLaserficheContentTypes();
        this.GetItemIdByTitle().then((results: IListItem[]) => {
            this.setState({ listItem: results });
            if (this.state.listItem != null) {
                const jsonValue = JSON.parse(this.state.listItem[0].JsonValue);
                if (jsonValue.length > 0) {
                    this.setState({
                        mappingRows: this.state.mappingRows.concat(jsonValue)
                    });
                }
            }
        });
        $('#sharePointValidationMapping').hide();
        $('#laserficheValidationMapping').hide();
        $('#validationOfMapping').hide();
    }

    //Get all laserfiche configuration created under manage configuration settings and append to Select element
    public GetAllLaserficheContentTypes() {
        let array = [];
        this.GetManageConfiguration().then((results: IListItem[]) => {
            if (results != null) {
                const jsonValue = JSON.parse(results[0].JsonValue);
                if (jsonValue.length > 0) {
                    for (var i = 0; i < jsonValue.length; i++) {
                        array.push({
                            name: jsonValue[i].ConfigurationName
                        });
                    }
                    this.setState({ laserficheContentTypes: array });
                }
            }
        });
    }

    //Get Manage Configurations setting value from the SharePoint list
    public async GetManageConfiguration(): Promise<IListItem[]> {
        let data: IListItem[] = [];
        let restApiUrl: string = this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('AdminConfigurationList')/Items?$select=Id,Title,JsonValue&$filter=Title eq 'ManageConfigurations'";
        try {
            const res = await fetch(restApiUrl, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json'
                },
            });
            const results = await res.json();
            if (results.value.length > 0) {
                for (var i = 0; i < results.value.length; i++) {
                    data.push(results.value[i]);
                }
                return data;
            }
            else {
                return null;
            }
        }
        catch (error) {
            console.log("error occured" + error);
        }
    }

    //Get all content types from SharePoint append to Select element
    public async GetAllSharePointContentTypes() {
        let array = [];
        let restApiUrl: string = this.props.context.pageContext.web.absoluteUrl + "/_api/web/contenttypes";
        try {
            const res = await fetch(restApiUrl, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json'
                },
            });
            const results = await res.json();
            for (var i = 0; i < results.value.length; i++) {
                array.push({
                    name: results.value[i].Name
                });
            }
            array.sort((a, b) => (a.name > b.name) ? 1 : -1);
            this.setState({ sharePointContentTypes: array });
        }
        catch (error) {
            console.log("error occured" + error);
        }
    }

    //If no mapping in list then create a New mapping in list or update the existing mapping in list
    public CreateNewMapping(idx, rows) {
        $('#sharePointValidationMapping').hide();
        $('#laserficheValidationMapping').hide();
        if (rows[idx].SharePointContentType == "Select") {
            $('#sharePointValidationMapping').show();
        }
        else if (rows[idx].LaserficheContentType == "Select") {
            $('#laserficheValidationMapping').show();
        }
        else {
            $('#sharePointValidationMapping').hide();
            $('#laserficheValidationMapping').hide();
            $('#validationOfMapping').hide();
            this.GetItemIdByTitle().then((results: IListItem[]) => {
                this.setState({ listItem: results });
                if (this.state.listItem != null) {
                    var entryExists = false;
                    let itemId = this.state.listItem[0].Id;
                    const jsonValue = JSON.parse(this.state.listItem[0].JsonValue);
                    if (jsonValue.length > 0) {
                        for (var i = 0; i < jsonValue.length; i++) {
                            if (jsonValue[i].id == rows[idx].id) {
                                entryExists = true;
                                break;
                            }
                        }
                        if (entryExists == true) {
                            this.UpdateExistingMapping(jsonValue, rows, idx, itemId);
                        }
                        else {
                            this.AddNewInExistingMapping(jsonValue, rows, idx, itemId);
                        }
                    }
                    else {
                        let restApiUrl: string = this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('AdminConfigurationList')/items(" + itemId + ")";
                        const row = [...this.state.mappingRows];
                        let newjsonValue = [{
                            "id": row[idx].id,
                            "SharePointContentType": row[idx].SharePointContentType,
                            "LaserficheContentType": row[idx].LaserficheContentType,
                            "toggle": true
                        }];
                        const jsonObject = JSON.stringify(newjsonValue);
                        const body: string = JSON.stringify({ 'Title': 'ManageMapping', 'JsonValue': jsonObject });
                        const options: ISPHttpClientOptions = {
                            headers: {
                                "Accept": "application/json;odata=nometadata",
                                "content-type": "application/json;odata=nometadata",
                                "odata-version": "",
                                'IF-MATCH': '*',
                                'X-HTTP-Method': 'MERGE'
                            },
                            body: body,
                        };
                        this.props.context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options);
                        rows[idx].toggle = !rows[idx].toggle;
                        this.setState({ mappingRows: rows });
                    }
                }
                else {
                    let restApiUrl: string = this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('AdminConfigurationList')/items";
                    const newRow = [...this.state.mappingRows];
                    let jsonValues = [{
                        "id": newRow[idx].id,
                        "SharePointContentType": newRow[idx].SharePointContentType,
                        "LaserficheContentType": newRow[idx].LaserficheContentType,
                        "toggle": true
                    }];
                    const jsonObject = JSON.stringify(jsonValues);
                    const body: string = JSON.stringify({ 'Title': 'ManageMapping', 'JsonValue': jsonObject });
                    const options: ISPHttpClientOptions = {
                        headers: {
                            "Accept": "application/json;odata=nometadata",
                            "content-type": "application/json;odata=nometadata",
                            "odata-version": "",
                        },
                        body: body,
                    };
                    this.props.context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options);
                    rows[idx].toggle = !rows[idx].toggle;
                    this.setState({ mappingRows: rows });
                }
            });
        }
    }

    //Add New mapping in existing json value
    public AddNewInExistingMapping(jsonValue, rows, idx, itemId) {
        var exitEntry = false;
        for (var i = 0; i < jsonValue.length; i++) {
            if (jsonValue[i].SharePointContentType == rows[idx].SharePointContentType) {
                exitEntry = true;
                break;
            }
        }
        if (exitEntry == false) {
            let restApiUrl: string = this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('AdminConfigurationList')/items(" + itemId + ")";
            const newJsonValue = [...jsonValue, {
                "id": rows[idx].id,
                "SharePointContentType": rows[idx].SharePointContentType,
                "LaserficheContentType": rows[idx].LaserficheContentType,
                "toggle": true
            }];
            const jsonObject = JSON.stringify(newJsonValue);
            const body: string = JSON.stringify({ 'Title': 'ManageMapping', 'JsonValue': jsonObject });
            const options: ISPHttpClientOptions = {
                headers: {
                    "Accept": "application/json;odata=nometadata",
                    "content-type": "application/json;odata=nometadata",
                    "odata-version": "",
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'MERGE'
                },
                body: body,
            };
            this.props.context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options);
            rows[idx].toggle = !rows[idx].toggle;
            this.setState({ mappingRows: rows });
            if(jsonValue.length+1==rows.length){
                $('#sharePointValidationMapping').hide();
                $('#laserficheValidationMapping').hide();
                $('#validationOfMapping').hide(); 
        }else{ 
          }
        }
        else {
            $('#validationOfMapping').show();
        }
    }

    //Update the existing mapping in json
    public UpdateExistingMapping(jsonValue, rows, idx, itemId) {
        var exitEntry = false;
        for (var i = 0; i < jsonValue.length; i++) {
            if (jsonValue[i].SharePointContentType == rows[idx].SharePointContentType) {
                if (jsonValue[i].id == rows[idx].id) {
                    exitEntry = false;
                    break;
                }
                else {
                    exitEntry = true;
                    break;
                }
            }
        }
        if (exitEntry == false) {
            for (var j = 0; j < jsonValue.length; j++) {
                if (jsonValue[j].id == rows[idx].id) {
                    jsonValue[j].SharePointContentType = rows[idx].SharePointContentType;
                    jsonValue[j].LaserficheContentType = rows[idx].LaserficheContentType;
                    jsonValue[j].toggle = !rows[idx].toggle;
                    break;
                }
            }
            let restApiUrl: string = this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('AdminConfigurationList')/items(" + itemId + ")";
            const newJsonValue = [...jsonValue];
            const jsonObject = JSON.stringify(newJsonValue);
            const body: string = JSON.stringify({ 'Title': 'ManageMapping', 'JsonValue': jsonObject });
            const options: ISPHttpClientOptions = {
                headers: {
                    "Accept": "application/json;odata=nometadata",
                    "content-type": "application/json;odata=nometadata",
                    "odata-version": "",
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'MERGE'
                },
                body: body,
            };
            this.props.context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options);
            rows[idx].toggle = !rows[idx].toggle;
            this.setState({ mappingRows: rows });
            if(rows.some(item =>item.SharePointContentType=="Select")){
                $('#sharePointValidationMapping').show();
            }else if(rows.some(item =>item.LaserficheContentType=="Select")){
                $('#laserficheValidationMapping').show();
            }else{
                $('#sharePointValidationMapping').hide();
                $('#laserficheValidationMapping').hide();
                $('#validationOfMapping').hide();
            }
        }
        else {
            $('#validationOfMapping').show();
        }
    }
    
    //Delete the mapping from the json
    public DeleteMapping(rows, idx) {
        this.GetItemIdByTitle().then((results: IListItem[]) => {
            this.setState({ listItem: results });
            if (this.state.listItem != null) {
                let itemId = this.state.listItem[0].Id;
                const jsonValue = JSON.parse(this.state.listItem[0].JsonValue);
                var entryExists = false;
                for (var i = 0; i < jsonValue.length; i++) {
                    if (jsonValue[i].id == rows[idx].id) {
                        entryExists = true;
                        break;
                    }
                }
                if (entryExists == true) {
                    jsonValue.splice(i, 1);
                    let restApiUrl: string = this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('AdminConfigurationList')/items(" + itemId + ")";
                    const newJsonValue = [...jsonValue];
                    const jsonObject = JSON.stringify(newJsonValue);
                    const body: string = JSON.stringify({ 'Title': 'ManageMapping', 'JsonValue': jsonObject });
                    const options: ISPHttpClientOptions = {
                        headers: {
                            "Accept": "application/json;odata=nometadata",
                            "content-type": "application/json;odata=nometadata",
                            "odata-version": "",
                            'IF-MATCH': '*',
                            'X-HTTP-Method': 'MERGE'
                        },
                        body: body,
                    };
                    this.props.context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options);
                    for(let q=jsonValue.length+1; q<rows.length; q++){
                        if(rows[idx].SharePointContentType==rows[q].SharePointContentType){
                            $('#validationOfMapping').hide(); 
                            break;
                        }else{
                            $('#validationOfMapping').show();
                        }
                    }
                }else{
                    if(jsonValue.length+1==rows.length){
                        $('#sharePointValidationMapping').hide();
                        $('#laserficheValidationMapping').hide();
                        $('#validationOfMapping').hide(); 
                }else{
                    for(let j=jsonValue.length; j<rows.length; j++){
                        if(rows[j].SharePointContentType=="Select"){
                            $('#sharePointValidationMapping').show();
                            break;
                        }else if(rows[j].LaserficheContentType=="Select"){
                            $('#laserficheValidationMapping').show();
                        }
                    }
                  }
                }
             }
        });
    }

    //Get ManageMapping value from the SharePoint list based on Title
    public async GetItemIdByTitle(): Promise<IListItem[]> {
        let array: IListItem[] = [];
        let restApiUrl: string = this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('AdminConfigurationList')/Items?$select=Id,Title,JsonValue&$filter=Title eq 'ManageMapping'";
        try {
            const res = await fetch(restApiUrl, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json'
                },
            });
            const results = await res.json();
            if (results.value.length > 0) {
                for (var i = 0; i < results.value.length; i++) {
                    array.push(results.value[i]);
                }
                return array;
            }
            else {
                return null;
            }
        }
        catch (error) {
            console.log("error occured" + error);
        }
    }

    //Add New Mapping in the UI
    public AddNewMapping = () => {
        var id = (+ new Date() + Math.floor(Math.random() * 999999)).toString(36);
        const item = {
            id: id,
            SharePointContentType: "Select",
            LaserficheContentType: "Select",
            toggle: false
        };
        this.setState({
            mappingRows: [...this.state.mappingRows, item]
        });

    }
    //Remove specific mapping from the UI
    public RemoveSpecificMapping = idx => () => {
        $('#deleteModal').data('id', idx);
        const rows = [...this.state.mappingRows];
        this.setState({ deleteSharePointcontentType: rows[idx].SharePointContentType });
        this.setState(() => { return { showDeleteModal: true }; });
    }
    //Getting confirmation in modal dialog to remove mapping
    public RemoveRow() {
        var id = $('#deleteModal').data('id');
        const rows = [...this.state.mappingRows];
        const deleteRows = [...this.state.mappingRows];
        rows.splice(id, 1);
        this.setState({ mappingRows: rows });
        this.DeleteMapping(deleteRows, id);
        this.setState(() => { return { showDeleteModal: false }; });
    }
    //Edit the specific mapping from the row
    public EditSpecificMapping = (idx) => () => {
        const rows = [...this.state.mappingRows];
        rows[idx].toggle = !rows[idx].toggle;
        this.setState({ mappingRows: rows });
    }
    //Save specific mapping from row
    public SaveSpecificMapping = (idx) => () => {
        const rows = [...this.state.mappingRows];
        this.CreateNewMapping(idx, rows);
    }
    //change event on Select elements
    public handleChange = idx => e => {
        var item = {
            id: e.target.id,
            name: e.target.name,
            value: e.target.value
        };
        const newRows = [...this.state.mappingRows];
        if (item.name == "SharePointContentType") {
            newRows[idx].SharePointContentType = item.value;
        }
        else if (item.name == "LaserficheContentType") {
            newRows[idx].LaserficheContentType = item.value;
        }
        this.setState({ mappingRows: newRows });
    }

    //Close the delete modal dialog
    public CloseModalUp() {
        this.setState(() => { return { showDeleteModal: false }; });
    }

    // Reset all the Recent edits to original form 
    public Reset() {
        this.setState(() => { return { showDeleteModal: false }; });
        this.GetAllSharePointContentTypes();
        this.GetAllLaserficheContentTypes();
        this.GetItemIdByTitle().then((results: IListItem[]) => {
            this.setState({ listItem: results });
            if (this.state.listItem != null) {
                const jsonValue = JSON.parse(this.state.listItem[0].JsonValue);
                if (jsonValue.length > 0) {
                    this.setState({
                        mappingRows: jsonValue
                    });
                }
            }
        });
        $('#sharePointValidationMapping').hide();
        $('#laserficheValidationMapping').hide();
        $('#validationOfMapping').hide();
    }
    
    //dynamic render the mapping and create table row elements
    public renderTableData() {
        let SharePointContents = this.state.sharePointContentTypes.map(v => (
            <option value={v.name}>{v.name}</option>
        ));
        let LaserficheContents = this.state.laserficheContentTypes.map(v => (
            <option value={v.name}>{v.name}</option>
        ));
        return this.state.mappingRows.map((item, index) => {
            if (item.toggle) {
                return (
                    <tr id="addr0" key={index}>
                        <td>
                            <select name="SharePointContentType" disabled className="custom-select" value={this.state.mappingRows[index].SharePointContentType} id={this.state.mappingRows[index].id} onChange={this.handleChange(index)}>
                                <option>Select</option>
                                {SharePointContents}
                            </select>
                        </td>
                        <td>
                            <select name="LaserficheContentType" disabled className="custom-select" value={this.state.mappingRows[index].LaserficheContentType} id={this.state.mappingRows[index].id} onChange={this.handleChange(index)}>
                                <option>Select</option>
                                {LaserficheContents}
                            </select>
                        </td>
                        <td className="text-center">
                            <a href="javascript:;" className="ml-3" onClick={this.EditSpecificMapping(index)}><span className="fa fa-pencil-alt fa-lg mt-2"></span></a>
                            <a href="javascript:;" className="ml-3" onClick={this.RemoveSpecificMapping(index)}><span className="fa fa-trash-alt fa-lg mt-2"></span></a>
                        </td>
                    </tr>
                );
            }
            else {
                return (
                    <tr id="addr0" key={index}>
                        <td>
                            <select name="SharePointContentType" className="custom-select" value={this.state.mappingRows[index].SharePointContentType} id={this.state.mappingRows[index].id} onChange={this.handleChange(index)}>
                                <option>Select</option>
                                {SharePointContents}
                            </select>
                        </td>
                        <td>
                            <select name="LaserficheContentType" className="custom-select" value={this.state.mappingRows[index].LaserficheContentType} id={this.state.mappingRows[index].id} onChange={this.handleChange(index)}>
                                <option>Select</option>
                                {LaserficheContents}
                            </select>
                        </td>
                        <td className="text-center">
                            <a href="javascript:;" className="ml-3" onClick={this.SaveSpecificMapping(index)} ><span className="fa fa-save fa-lg mt-2"></span></a>
                            <a href="javascript:;" className="ml-3" onClick={this.RemoveSpecificMapping(index)}><span className="fa fa-trash-alt fa-lg mt-2"></span></a>
                        </td>
                    </tr>
                );
            }
        });
    }
    public render(): React.ReactElement {
        let viewSharePointContentTypes = this.props.context.pageContext.web.absoluteUrl + "/_layouts/15/mngctype.aspx";
        return (
            <div className="">
                <div className="">
                    <div className="container-fluid p-3" style={{"maxWidth":"85%","marginLeft":"-26px"}}>

                        <div className="p-3">
                            <div className="card rounded-0">
                                <div className="card-header d-flex justify-content-between">
                                    <div>
                                        <h6 className="mb-0">Content Type Mappings Laserfiche</h6>
                                    </div>
                                    <div>
                                        <a href="" onClick={() => window.open(viewSharePointContentTypes)} target="_blank">View SharePoint Content Types</a>
                                    </div>
                                </div>
                                <div className="card-body">
                                    <table className="table table-sm">
                                        <thead>
                                            <tr>
                                                <th className="text-center" style={{ width: "45%" }}>SharePoint Content Type</th>
                                                <th className="text-center" style={{ width: "45%" }}>Laserfiche Profile</th>
                                                <th className="text-center">Action</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {this.renderTableData()}
                                        </tbody>
                                    </table>
                                </div>
                                <div id="sharePointValidationMapping" style={{ color: "red" }}><span>Please select a content type from the SharePoint Content Type drop down</span></div>
                                <div id="laserficheValidationMapping" style={{ color: "red" }}><span>Please select a content type from the Laserfiche Profile drop down</span></div>
                                <div id="validationOfMapping" style={{ color: "red" }}><span>Already Mapping exists for this SharePoint content type</span></div>
                                <div className="card-footer bg-transparent">
                                    <a  className="btn btn-primary pl-5 pr-5 float-right" style={{"marginLeft":"10px"}} onClick={this.Reset}>Reset</a>
                                    <a href="javascript:;" className="btn btn-primary pl-5 pr-5 float-right" onClick={this.AddNewMapping}>Add</a>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div className="modal" id="deleteModal" hidden={!this.state.showDeleteModal} data-backdrop="static" data-keyboard='false'>
                        <div className="modal-dialog modal-dialog-centered">
                            <div className="modal-content">
                                <div className="modal-header">
                                    <h5 className="modal-title" id="ModalLabel">Delete Confirmation</h5>
                                    <button type="button" className="close" data-dismiss="modal" aria-label="Close" onClick={() => this.CloseModalUp()}>
                                        <span aria-hidden="true">&times;</span>
                                    </button>
                                </div>
                                <div className="modal-body">
                                    Do you want to permanently delete the "{this.state.deleteSharePointcontentType}" mapping?{/*  for the "{this.state.deleteSharePointcontentType}" Content Type? */}
                                </div>
                                <div className="modal-footer">
                                    <button type="button" className="btn btn-primary btn-sm" data-dismiss="modal" onClick={() => this.RemoveRow()}>OK</button>
                                    <button type="button" className="btn btn-secondary btn-sm" data-dismiss="modal" onClick={() => this.CloseModalUp()}>Cancel</button>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        );
    }
}


