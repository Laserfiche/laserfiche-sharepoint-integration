import * as React from "react";
import * as $ from 'jquery';
import { NavLink } from 'react-router-dom';
import { IManageConfigurationPageProps } from './IManageConfigurationPageProps';
import { IManageConfigurationPageState } from './IManageConfigurationPageState';
import { IListItem } from './IListItem';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
require('../../../../Assets/CSS/bootstrap.min.css');
require('../../../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');
require('../../../../Assets/CSS/adminConfig.css');
require('../../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');

export default class ManageConfigurationsPage extends React.Component<IManageConfigurationPageProps, IManageConfigurationPageState> {
    constructor(props: IManageConfigurationPageProps) {
        super(props);
        this.state = {
            configurationRows: [],
            listItem: [],
            showDeleteModal: false,
            configurationName: ''
        };
    }
    //On load get list of configurations created from the Admin Configuration list
    public componentDidMount(): void {
        this.setState(() => { return { showDeleteModal: false }; });
        this.GetItemIdByTitle().then((results: IListItem[]) => {
            this.setState({ listItem: results });
            if (this.state.listItem != null) {
                const jsonValue = JSON.parse(this.state.listItem[0].JsonValue);
                if (jsonValue.length > 0) {
                    this.setState({
                        configurationRows: this.state.configurationRows.concat(jsonValue)
                    });
                }
            }
        });
    }
    //Get items from the AdminConfiguratiion list based on Title 'ManageConfiguration'
    public async GetItemIdByTitle(): Promise<IListItem[]> {
        let array: IListItem[] = [];
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

    //Remove specific configuration from the list
    public RemoveSpecificConfiguration = idx => () => {
        $('#deleteModal').data('id', idx);
        const rows = [...this.state.configurationRows];
        this.setState({ configurationName: rows[idx].ConfigurationName });
        this.setState(() => { return { showDeleteModal: true }; });
    }

    //Remove row on click on delete button
    public RemoveRow() {
        var id = $('#deleteModal').data('id');
        const rows = [...this.state.configurationRows];
        const deleteRows = [...this.state.configurationRows];
        rows.splice(id, 1);
        this.setState({ configurationRows: rows });
        this.DeleteMapping(deleteRows, id);
        this.setState(() => { return { showDeleteModal: false }; });
    }

    //Close Modal dialog box
    public CloseModalUp() {
        this.setState(() => { return { showDeleteModal: false }; });
    }

    //Delete the selected configuration from the list 
    public DeleteMapping(rows, idx) {
        this.GetItemIdByTitle().then((results: IListItem[]) => {
            this.setState({ listItem: results });
            if (this.state.listItem != null) {
                let itemId = this.state.listItem[0].Id;
                const jsonValue = JSON.parse(this.state.listItem[0].JsonValue);
                for (var i = 0; i < jsonValue.length; i++) {
                    if (jsonValue[i].ConfigurationName == rows[idx].ConfigurationName) {
                        jsonValue.splice(i, 1);
                        let restApiUrl: string = this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('AdminConfigurationList')/items(" + itemId + ")";
                        const newJsonValue = [...jsonValue];
                        const jsonObject = JSON.stringify(newJsonValue);
                        const body: string = JSON.stringify({ 'Title': 'ManageConfigurations', 'JsonValue': jsonObject });
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
                        break;
                    }
                }
            }
        });
    }

    //Dynamically render list of configurations created in the table format 
    public renderTableData() {
        return this.state.configurationRows.map((item, index) => {
            return (
                <tr id="addr0" key={index}>
                    <td>{this.state.configurationRows[index].ConfigurationName}</td>
                    <td className="text-center">
                        <span><NavLink to={'/EditManageConfiguration/' + this.state.configurationRows[index].ConfigurationName} style={{ marginRight: "18px", fontWeight: '500', fontSize: '15px' }}><span className="fa fa-pencil-alt fa-lg"></span></NavLink></span>
                        <a href="javascript:;" className="ml-3" onClick={this.RemoveSpecificConfiguration(index)}><span className="fa fa-trash-alt fa-lg"></span></a>
                    </td>
                </tr>
            );
        });
    }
    
    public render(): React.ReactElement {
        return (
            <div>
                <div className="container-fluid p-3" style={{"maxWidth":"85%","marginLeft":"-26px"}}>
                    <main className="bg-white shadow-sm">
                        <div className="p-3">
                            <div className="card rounded-0">
                                <div className="card-header d-flex justify-content-between pt-1 pb-1">
                                    <div>
                                    </div>
                                    <div>
                                        <NavLink to="/AddNewManageConfiguration" style={{ marginRight: "18px", fontWeight: '500', fontSize: '15px' }}><a className="btn btn-primary pl-5 pr-5">Add Profile</a></NavLink>
                                    </div>
                                </div>
                                <div className="card-body">
                                    <table className="table table-bordered table-striped table-hover">
                                        <thead>
                                            <tr>
                                                <th className="text-center">Profile Name</th>
                                                <th className="text-center" style={{"width":"30%"}}>Action</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {this.renderTableData()}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    </main>
                </div>
                <div>
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
                                    Do you want to permanently delete "{this.state.configurationName}"?
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