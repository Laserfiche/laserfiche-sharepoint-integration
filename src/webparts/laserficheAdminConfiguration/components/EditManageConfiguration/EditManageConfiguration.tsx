import * as React from "react";
import * as $ from 'jquery';
import * as bootstrap from 'bootstrap';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { NavLink } from 'react-router-dom';
import { IEditManageConfigurationProps } from './IEditManageConfigurationProps';
import { IEditManageConfigurationState } from './IEditManageConfigurationState';
import { IListItem } from "./IListItem";
import { Spinner, SpinnerSize } from "office-ui-fabric-react";
import { ODataValueContextOfIListOfWTemplateInfo, ODataValueOfIListOfTemplateFieldInfo, RepositoryApiClient, WTemplateInfo, EntryType,Shortcut } from "@laserfiche/lf-repository-api-client";
import {LfFieldsService,LfRepoTreeNode, LfRepoTreeNodeService} from '@laserfiche/lf-ui-components-services';
import { LoginState, TreeNode } from "@laserfiche/types-lf-ui-components";
import { IRepositoryApiClientExInternal } from "../../../../repository-client/repository-client-types";
import { RepositoryClientExInternal } from "../../../../repository-client/repository-client";
import { clientId } from "../../../constants";
require('../../../../Assets/CSS/bootstrap.min.css');
require('@fortawesome/fontawesome-free/css/all.min.css');
require('../../../../Assets/CSS/adminConfig.css');
require('../../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');

declare global {
  namespace JSX {
    interface IntrinsicElements {
      ['lf-login']: any;
      ['lf-repository-browser']: any;
    }
  } 
}

export default class EditManageConfiguration extends React.Component<IEditManageConfigurationProps, IEditManageConfigurationState> {
  public loginComponent: React.RefObject<any>;
  public repositoryBrowser: React.RefObject<any>;
  public divRef: React.RefObject<HTMLDivElement>;
  public repoClient: IRepositoryApiClientExInternal;
  public lfRepoTreeService: LfRepoTreeNodeService;
  public lfFieldsService: LfFieldsService;
  public showTree: boolean = false;
  public entrySelected: LfRepoTreeNode | undefined;

  constructor(props: IEditManageConfigurationProps) {
    super(props);
    SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/indigo-pink.css');
    SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ms-office-lite.css');
    this.loginComponent = React.createRef();
    this.repositoryBrowser = React.createRef();
    this.divRef = React.createRef();
    this.state = {
      mappingList: [],
      listItem: [],
      laserficheTemplates: [],
      sharePointFields: [],
      laserficheFields: [],
      documentNames: [],
      loadingContent: false,
      hideContent: true,
      showFolderModal: false,
      showtokensModal: false,
      showDeleteModal: false,
      showConfirmModal: false,
      lfSelectedFolder:{
        //selectedNodeUrl: '', 
        selectedFolderPath: '', 
        //selectedFolderName: ''
      },
      shouldShowOpen: false, 
      shouldShowSelect: false,
      shouldDisableSelect: false,
      region: this.props.devMode ? 'a.clouddev.laserfiche.com' : 'laserfiche.com'
    };
  }
  
  //On component load get content types from SharePoint,laserfiche templates and existing configuration from the list
  public async componentDidMount(): Promise<void> {
    this.setState({ hideContent: true });
    this.setState({ loadingContent: false });
    const configurationName = this.props.match.params.name;
    await SPComponentLoader.loadScript('https://cdn.jsdelivr.net/npm/zone.js@0.11.4/bundles/zone.umd.min.js');
    await SPComponentLoader.loadScript('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ui-components.js');
    SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/indigo-pink.css');
    SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ms-office-lite.css');

    $('#tokens').css("margin-top", "4px !important");
    $('[data-toggle="tooltip"]').tooltip({
      html: true,
      placement: "right",
      title: "<div class='toolTipCustom'>Replace - SharePoint file replaced with Link to Laserfiche File</div><div class='toolTipCustom'>Copy - Keep File in Laserfiche and SharePoint</div><div class='toolTipCustom'>Move and Delete - SharePoint file deleted after Import to Laserfiche</div>"
    });
    //Add event listener to lf login and folder browser component
    this.loginComponent.current.addEventListener('loginCompleted', this.loginCompleted);
    this.loginComponent.current.addEventListener('logoutCompleted', this.logoutCompleted);
    //this.folderbrowser.current.addEventListener('okClick', this.onOkClick);
    //this.folderbrowser.current.addEventListener('cancelClick', this.onCancelClick);
    this.setState(() => { return { showFolderModal: false, showtokensModal: false, showDeleteModal: false, showConfirmModal: false }; });
    $('#sharePointFieldMapping').hide();
    $('#laserficheFieldMapping').hide();
    $('#addMapping').hide();
    //Checking the login state of lf login component
    const loggedOut: boolean = this.loginComponent.current.state === LoginState.LoggedOut;
    if (loggedOut) {
      $('.ManageConfigurationLink').hide();
      $('.ManageMappingLink').hide();
      $('.HomeLink').hide();
    }
    else {
      $('.ManageConfigurationLink').show();
      $('.ManageMappingLink').show();
      $('.HomeLink').show();
    }

    //Get access token and repoclient for api calls
    await this.getAndInitializeRepositoryClientAndServicesAsync();

    this.GetItemIdByTitle().then((results: IListItem[]) => {
      this.setState({ listItem: results });
      this.GetTemplateDefinitions().then((templates: string[]) => {
        templates.sort();
        this.setState({ laserficheTemplates: templates });
        if (this.state.listItem != null) {
          const jsonValue = JSON.parse(this.state.listItem[0].JsonValue);
          if (jsonValue.length > 0) {
            for (var i = 0; i < jsonValue.length; i++) {
              if (jsonValue[i].ConfigurationName == configurationName) {
                $('#configurationTitle').text(jsonValue[i].ConfigurationName);
                $('#documentName').val(jsonValue[i].DocumentName);
                $('select[id="documentTemplate"]').val(jsonValue[i].DocumentTemplate);
                $('#destinationPath').val(jsonValue[i].DestinationPath);
                $('#entryId').val(jsonValue[i].EntryId);
                $('select[id="action"]').val(jsonValue[i].Action);
                this.MappingFields(jsonValue[i].DocumentTemplate, jsonValue[i].SharePointFields, jsonValue[i].LaserficheFields);
              }
            }
          }
        }
      });
    });

    //Get all Site columns from the SharePoint site
    this.GetAllSharePointSiteColumns().then((contents: any) => {
      contents.sort((a, b) => (a.DisplayName > b.DisplayName) ? 1 : -1);
      this.setState({
        sharePointFields: this.state.sharePointFields.concat(contents)
      });
    });
    //Get document name from the Document Configuration list
    /* this.GetDocumentName().then((names: string[]) => {
      this.setState({ documentNames: names });
    }); */
  }

  //Okay function on folder selection in Folders modal dialog box
  public onOkClick = (ev: Event) => {
    const selectedNode = (ev as CustomEvent<TreeNode>).detail;
    $('#entryId').val(selectedNode.id);
    $('#destinationPath').val(selectedNode.path);
    this.divRef.current.innerText = "Selected Folder:" + selectedNode.name;
    this.setState(() => { return { showFolderModal: false }; });
  }

  //Cancel function on folders modal dialog box
  public onCancelClick = (ev: Event) => {
    this.setState(() => { return { showFolderModal: false }; });
  }

  //Laserfiche LF login component 
  public loginCompleted = async () => {
    await this.getAndInitializeRepositoryClientAndServicesAsync();
    $('.ManageConfigurationLink').show();
    $('.ManageMappingLink').show();
    $('.HomeLink').show();
  }
  public logoutCompleted = async () => {
    $('.ManageConfigurationLink').hide();
    $('.ManageMappingLink').hide();
    $('.HomeLink').hide();
    window.location.href = this.props.context.pageContext.web.absoluteUrl + this.props.laserficheRedirectPage;
  }

  private async getAndInitializeRepositoryClientAndServicesAsync() {
    const accessToken = this.loginComponent?.current?.authorization_credentials?.accessToken;
    if (accessToken) {

      await this.ensureRepoClientInitializedAsync();

      // create the tree service to interact with the LF Api
      this.lfRepoTreeService = new LfRepoTreeNodeService(this.repoClient!);
      // by default all entries are viewable
      this.lfRepoTreeService.viewableEntryTypes = [EntryType.Folder, EntryType.Shortcut];
      //await this.initializeTreeAsync();

    }
    else {
      // user is not logged in
    }
  }

  public async ensureRepoClientInitializedAsync(): Promise<void> {
    if (!this.repoClient) {
      const repoClientCreator = new RepositoryClientExInternal(this.loginComponent);
      this.repoClient = await repoClientCreator.createRepositoryClientAsync();
    }
  }

  //Getting folder browser component in Folder dialog box
  public async initializeTreeAsync() {
    /* this.showTree = true;
    await this.folderbrowser.current.initAsync({
      treeService: this.lfRepoTreeService
    }); */
    if (!this.repoClient) {
      throw new Error('RepoId is undefined');
    }
    this.repositoryBrowser.current?.addEventListener('entrySelected', this.onEntrySelected );
    let focusedNode: LfRepoTreeNode | undefined;
    if (this.state.lfSelectedFolder.selectedFolderPath != "") {
      const repoId = await this.repoClient.getCurrentRepoId();
      const focusedNodeByPath = await this.repoClient.entriesClient.getEntryByPath({
          repoId: repoId,
          fullPath: this.state?.lfSelectedFolder.selectedFolderPath
        });
      const repoName = await this.repoClient.getCurrentRepoName();
      const focusedNodeEntry = focusedNodeByPath?.entry;
      if (focusedNodeEntry) {
        focusedNode = this.lfRepoTreeService?.createLfRepoTreeNode(focusedNodeEntry, repoName);
      }
    }
    await this.repositoryBrowser?.current?.initAsync(this.lfRepoTreeService!, focusedNode);
  }

  public onSelectFolder = async () => {
    if (!this.repoClient) {
      throw new Error('Repo Client is undefined.');
    }
    if (!this.loginComponent.current?.account_endpoints) {
      throw new Error('LfLoginComponent is not found.');
    }
    const selectedNode = this.repositoryBrowser.current?.currentFolder as LfRepoTreeNode;
    let entryId = Number.parseInt(selectedNode.id, 10);
    const selectedFolderPath = selectedNode.path;
    $('#entryId').val(selectedNode.id);
    $('#destinationPath').val(selectedNode.path);
    if (selectedNode.entryType === EntryType.Shortcut) {
      if (selectedNode.targetId)
      entryId = selectedNode.targetId;
    }
    const repoId = (await this.repoClient.getCurrentRepoId());
    const waUrl = this.loginComponent.current.account_endpoints.webClientUrl;
    this.setState({
      lfSelectedFolder: {
        //selectedNodeUrl: getEntryWebAccessUrl(entryId.toString(), repoId, waUrl, selectedNode.isContainer) ?? '',
        //selectedFolderName: this.getFolderNameText(entryId, selectedFolderPath),
        selectedFolderPath: selectedFolderPath
      },
      shouldShowOpen: false,
      showFolderModal: false,
      shouldShowSelect: false,
    });
  }

   public onClickCancelButton = () => {
    this.setState({
      showFolderModal: false,
      shouldShowOpen: false,
      shouldShowSelect: false
    });
  }

  public getShouldShowSelect(): boolean {
    return !this.entrySelected && !!this.repositoryBrowser?.current?.currentFolder;
  }

  public getShouldShowOpen(): boolean {
    return !!this.entrySelected;
  }

  public getShouldDisableSelect(): boolean {
    return !this.isNodeSelectable(this.repositoryBrowser?.current?.currentFolder as LfRepoTreeNode);
  }

  public isNodeSelectable = (node: LfRepoTreeNode) => {
    if (node.entryType == EntryType.Folder) {
      return true;
    }
    else if (node.entryType == EntryType.Shortcut && node.targetType == EntryType.Folder) {
      return true;
    }
    else {
      return false;
    }
  }

  public onEntrySelected = (event: any) => {
    const treeNodesSelected: LfRepoTreeNode[] = event.detail;
    this.entrySelected = treeNodesSelected?.length > 0 ? treeNodesSelected[0] : undefined;
    this.setState({
      shouldShowOpen: this.getShouldShowOpen(),
      shouldShowSelect: this.getShouldShowSelect(),
      shouldDisableSelect: this.getShouldDisableSelect(),
    });
  }
  
  public folderCancelClick = () => {
    this.setState({ showFolderModal: false });
  }

  public onOpenNode = async () => {
    await this.repositoryBrowser?.current?.openSelectedNodesAsync();
    this.setState({
      shouldShowOpen:  this.getShouldShowOpen(),
      shouldShowSelect: this.getShouldShowSelect()
    });
  }

  //Getting document name from Document Configuration list
  public async GetDocumentName(): Promise<string[]> {
    let name: string[] = [];
    let restApiUrl: string = this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('DocumentNameConfigList')/Items?$select=Title";
    try {
      const res = await fetch(restApiUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
        },
      });
      const results = await res.json();
      for (var i = 0; i < results.value.length; i++) {
        name.push(results.value[i].Title);
      }
      return name;
    }
    catch (error) {
      console.log("error occured" + error);
    }
  }

  //Get all Site columns from SharePoint site 
  public async GetAllSharePointSiteColumns(): Promise<any> {
    let array = [];
    let restApiUrl: string = this.props.context.pageContext.web.absoluteUrl + "/_api/web/fields?$filter=(Hidden ne true and Group ne '_Hidden')";
    try {
      const res = await fetch(restApiUrl, {
        method: 'GET',
        headers: {
          'Accept': "application/json;odata=nometadata",
          "content-type": "application/json;odata=nometadata",
          "odata-version": "",
        },
      });
      const results = await res.json();
      if (results.value.length > 0) {
        for (var i = 0; i < results.value.length; i++) {
          array.push({ "DisplayName": results.value[i].Title + "[" + results.value[i].TypeAsString + "]", "InternalName": results.value[i].InternalName + "[" + results.value[i].TypeAsString + "]" });
        }
        return array;
      }
    }
    catch (error) {
      console.log("error occured" + error);
    }
  }

  //Get templates from Laserfiche
  public async GetTemplateDefinitions(): Promise<string[]> {
    let array = [];
    const templateInfo: WTemplateInfo[] = [];
    const repoId = await this.repoClient.getCurrentRepoId();
    await this.repoClient.templateDefinitionsClient.getTemplateDefinitionsForEach({
      callback: async (response: ODataValueContextOfIListOfWTemplateInfo) => {
        if(response.value) {
          templateInfo.push(...response.value);
        }
        return true;
      },
      repoId
    });
    array = templateInfo.map((value) => value.name);
    return array;
  }

  //Get items from SharePoint Admin Configuration list based on Title ManageConfiguration
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

  //Get laserfiche fields based on template change
  public OnChangeTemplate() {
    $('#sharePointFieldMapping').hide();
    $('#laserficheFieldMapping').hide();
    $('#addMapping').hide();
    let templatename = $("#documentTemplate option:selected").text();
    this.GetLaserficheFields(templatename).then((fields: string[]) => {
      if (fields != null) {
        this.setState({ laserficheFields: fields });
        $('#tableEditBodyId').show();
        let array = [];
        for (let index = 0; index < fields.length; index++) {
          var id = (+ new Date() + Math.floor(Math.random() * 999999)).toString(36);
          let laserficheField = fields[index]["InternalName"];
          if (laserficheField.indexOf("[Required:true]") != -1) {
            array.push({ "id": id, "SharePointField": "Select", "LaserficheField": fields[index]["InternalName"] });
          }
        }
        this.setState({ mappingList: array });
      }
      else {
        this.setState({ laserficheFields: [] });
        $('#tableEditBodyId').hide();
      }
      for(let j=0;j<this.state.mappingList.length;j++){
        var spanId='a'+j;
        document.getElementById(spanId).style.display='none';
      }
    });
  }

  //Get laserfiche fields based on template 
  public async GetLaserficheFields(templatename): Promise<string[]> {
    if (templatename != "None") {
      let array = [];
      const repoId = await this.repoClient.getCurrentRepoId();
      const apiTemplateResponse: ODataValueOfIListOfTemplateFieldInfo = await this.repoClient.templateDefinitionsClient.getTemplateFieldDefinitionsByTemplateName(
        { repoId, templateName: templatename }
      );
      const fieldsValues = apiTemplateResponse.value;
      for (var i = 0; i < fieldsValues.length; i++) {
        array.push({ "DisplayName": fieldsValues[i].name + "[" + fieldsValues[i].fieldType + "]", "InternalName": fieldsValues[i].name + "[" + fieldsValues[i].fieldType + "]" + "[" + "Required:" + fieldsValues[i].isRequired + "]" + "[" + "length:" + fieldsValues[i].length + "]" + "[" + "constraint:" + fieldsValues[i].constraint + "]" });
      }
      return array;
    }
    else {
      return null;
    }
  }

  // Creating mapping fields if Laserfiche field is deleted in repository or any parameter updated for the field
  public async MappingFields(DocumentTemplate, SharePointFields, LaserficheFields) {
    if (DocumentTemplate != "None") {
      let array = [];
      const repoId = await this.repoClient.getCurrentRepoId();
      const apiTemplateResponse: ODataValueOfIListOfTemplateFieldInfo = await this.repoClient.templateDefinitionsClient.getTemplateFieldDefinitionsByTemplateName(
        { repoId, templateName: DocumentTemplate }
      );
      const fieldsValues = apiTemplateResponse.value;
      for (var i = 0; i < fieldsValues.length; i++) {
        array.push({ "DisplayName": fieldsValues[i].name + "[" + fieldsValues[i].fieldType + "]", "InternalName": fieldsValues[i].name + "[" + fieldsValues[i].fieldType + "]" + "[" + "Required:" + fieldsValues[i].isRequired + "]" + "[" + "length:" + fieldsValues[i].length + "]" + "[" + "constraint:" + fieldsValues[i].constraint + "]" });
      }
      this.setState({ laserficheFields: array });
      let array1 = [];
      for (let index = 0; index < SharePointFields.length; index++) {
        var id = (+ new Date() + Math.floor(Math.random() * 999999)).toString(36);
        if (DocumentTemplate != "None") {
          let laserfciheItems = array;
          let laserficheValue = LaserficheFields[index].split("[")[0];
          let laserficheInternalName = "Select";
          laserfciheItems.some(item => {
            let fieldName = item.DisplayName.split("[")[0];
            if (fieldName === laserficheValue) {
              laserficheInternalName = item.InternalName;
            }
          });
          array1.push({ "id": id, "SharePointField": SharePointFields[index], "LaserficheField": laserficheInternalName });
        }
      }
      this.setState({ mappingList: array1 });
      const rows = [...this.state.mappingList];
      let missingFields = [];
      for (let j = 0; j < rows.length; j++) {
        missingFields.push(rows[j].LaserficheField.split("[")[0]);
      }
      let newArray = [];
      for (var p = 0; p < array.length; p++) {
        let laserficheField = array[p]["InternalName"].split("[")[0];
        if (array[p]["InternalName"].indexOf("[Required:true]") != -1) {
          if (missingFields.indexOf(laserficheField) != -1) {

          }
          else {
            var id1 = (+ new Date() + Math.floor(Math.random() * 999999)).toString(36);
            newArray.push({ "id": id1, "SharePointField": "Select", "LaserficheField": array[p]["InternalName"] });
          }
        }
      }
      this.setState({
        mappingList: this.state.mappingList.concat(newArray)
      });
    }
    else {
      this.setState({ laserficheFields: [] });
    }
    this.setState({ loadingContent: true });
    this.setState({ hideContent: false });
  }

  //Add new mapping fields
  public AddNewMappingFields() {
    $('#sharePointFieldMapping').hide();
    $('#laserficheFieldMapping').hide();
    $('#addMapping').hide();
    let templatename = $("#documentTemplate option:selected").text();
    if (templatename != "None") {
      var id = (+ new Date() + Math.floor(Math.random() * 999999)).toString(36);
      const item = {
        id: id,
        SharePointField: "Select",
        LaserficheField: "Select",
      };
      this.setState({
        mappingList: [...this.state.mappingList, item]
      });
    }
    else {
      $('#addMapping').show();
    }
  }

  //OnChange functionality on elemnts
  public handleChange = idx => e => {
    var rowID;
    var item = {
      id: e.target.id,
      name: e.target.name,
      value: e.target.value
    };
    var rowsArray = this.state.mappingList;
    var newRow = rowsArray.map((row, i) => {
      for (var key in row) {
        if (key == item.name && row.id == item.id) {
          row[key] = item.value;
        }
      }
      return row;
    });
    this.setState({ mappingList: newRow });
    var rows = [...this.state.mappingList];
      for(let j=0; j<rows.length; j++){
        var spanId='a'+j;
        document.getElementById(spanId).style.display='none';
      }
      for(let i=0; i<rows.length; i++){
        if(rows[i]["SharePointField"].includes('[') && rows[i]["LaserficheField"].includes('[') ){
        var sharepointfieldtype=rows[i]["SharePointField"].split('[')[1];
        var spFieldtype=sharepointfieldtype.slice(0,-1);
        var laserfichepointfieldtype=rows[i]["LaserficheField"].split('[')[1];
        var lfFieldtype=laserfichepointfieldtype.slice(0,-1);
        rowID='a'+i;
        if(lfFieldtype=="DateTime"||lfFieldtype=="Date"||lfFieldtype=="Time"){
          if(spFieldtype!="DateTime"){
            document.getElementById(rowID).style.display="inline-block";
            document.getElementById(rowID).title=`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${lfFieldtype}`;
          }
        }else if(lfFieldtype=="LongInteger" ||lfFieldtype=="ShortInteger"){
          if(spFieldtype!="Number"){
            document.getElementById(rowID).style.display="inline-block";
            document.getElementById(rowID).title=`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${lfFieldtype}`;
          }
        }else if(lfFieldtype=="Number"){
          if(spFieldtype !="Number" && spFieldtype !="Currency"){
            document.getElementById(rowID).style.display="inline-block";
            document.getElementById(rowID).title=`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${lfFieldtype}`;
          }
      } else if(lfFieldtype=="List"){
        if(spFieldtype!="Choice"){
          document.getElementById(rowID).style.display="inline-block";
          document.getElementById(rowID).title=`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${lfFieldtype}`;
        }
    } 
    }
    }
  }
  //Remove specfic fields mapping
  public RemoveSpecificMapping = idx => () => {
    $('#sharePointFieldMapping').hide();
    $('#laserficheFieldMapping').hide();
    $('#addMapping').hide();
    $('#deleteModal').data('id', idx);
    this.setState(() => { return { showDeleteModal: true }; });
  }
  //delete the SharePoint and Laserfiche fields mapping
  public DeleteMapping() {
    var id = $('#deleteModal').data('id');
    const rows = [...this.state.mappingList];
    rows.splice(id, 1);
    this.setState({ mappingList: rows });
    this.setState(() => { return { showDeleteModal: false }; });
    for(let i=0; i<rows.length; i++){
      var spanId='a'+i;
      document.getElementById(spanId).style.display='none';
    }
    for(let i=0; i<rows.length; i++){
      var sharepointfieldtype=rows[i]["SharePointField"].split('[')[1];
      var spFieldtype=sharepointfieldtype.slice(0,-1);
      var laserfichepointfieldtype=rows[i]["LaserficheField"].split('[')[1];
      var lfFieldtype=laserfichepointfieldtype.slice(0,-1);
      var rowID='a'+i;
      if(lfFieldtype=="DateTime"||lfFieldtype=="Date"||lfFieldtype=="Time"){
        if(spFieldtype!="DateTime"){
          document.getElementById(rowID).style.display="inline-block";
          document.getElementById(rowID).title=`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${lfFieldtype}`;
        }
      }else if(lfFieldtype=="LongInteger" ||lfFieldtype=="ShortInteger"){
        if(spFieldtype!="Number"){
          document.getElementById(rowID).style.display="inline-block";
          document.getElementById(rowID).title=`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${lfFieldtype}`;
        }
      }else if(lfFieldtype=="Number"){
        if(spFieldtype !="Number" && spFieldtype !="Currency"){
          document.getElementById(rowID).style.display="inline-block";
          document.getElementById(rowID).title=`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${lfFieldtype}`;
        }
    } else if(lfFieldtype=="List"){
      if(spFieldtype!="Choice"){
        document.getElementById(rowID).style.display="inline-block";
        document.getElementById(rowID).title=`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${lfFieldtype}`;
      }
  } 
  }
  }

  //Close the Delete Modal pop up 
  public CloseModalUp() {
    this.setState(() => { return { showDeleteModal: false }; });
  }

  //Update the existing configuration in SharePoint list
  public SaveNewManageConfigurtaion() {
    var rowID;
    $('#sharePointFieldMapping').hide();
    $('#laserficheFieldMapping').hide();
    $('#addMapping').hide();
    let validation: boolean = true;
    const rows = [...this.state.mappingList];
    if (rows.some(item => item.SharePointField === "Select") && $("#documentTemplate option:selected").text() != "None") {
      $('#sharePointFieldMapping').show();
    }
    else if (rows.some(items => items.LaserficheField === "Select") && $("#documentTemplate option:selected").text() != "None") {
      $('#laserficheFieldMapping').show();
    }
    else {
      for(let j=0; j<rows.length; j++){
        var spanId='a'+j;
        document.getElementById(spanId).style.display='none';
      }
      for(let i=0; i<rows.length; i++){
        var sharepointfieldtype=rows[i]["SharePointField"].split('[')[1];
        var spFieldtype=sharepointfieldtype.slice(0,-1);
        var laserfichepointfieldtype=rows[i]["LaserficheField"].split('[')[1];
        var lfFieldtype=laserfichepointfieldtype.slice(0,-1);
        rowID='a'+i;
        if(lfFieldtype=="DateTime"||lfFieldtype=="Date"||lfFieldtype=="Time"){
          if(spFieldtype!="DateTime"){
            validation=false;
            document.getElementById(rowID).style.display="inline-block";
            document.getElementById(rowID).title=`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${lfFieldtype}`;
          }
        }else if(lfFieldtype=="LongInteger" ||lfFieldtype=="ShortInteger"){
          if(spFieldtype!="Number"){
            validation=false;
            document.getElementById(rowID).style.display="inline-block";
            document.getElementById(rowID).title=`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${lfFieldtype}`;
          }
        }else if(lfFieldtype=="Number"){
          if(spFieldtype !="Number" && spFieldtype !="Currency"){
            validation=false;
            document.getElementById(rowID).style.display="inline-block";
            document.getElementById(rowID).title=`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${lfFieldtype}`;
          }
      } else if(lfFieldtype=="List"){
        if(spFieldtype!="Choice"){
          validation=false;
          document.getElementById(rowID).style.display="inline-block";
          document.getElementById(rowID).title=`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${lfFieldtype}`;
        }
    } 
    }
    if(validation){
      $('#sharePointFieldMapping').hide();
      $('#laserficheFieldMapping').hide();
      $('#addMapping').hide();
      var documentName = document.getElementById('documentName')["value"];
      var docTemp = document.getElementById('documentTemplate')["value"];
      var destPath = document.getElementById('destinationPath')["value"];
      var entryId = document.getElementById('entryId')["value"];
      var action = document.getElementById('action')["value"];
      let sharepointFields = [];
      let laserficheFields = [];
      if (docTemp != "None") {
        for (let i = 0; i < rows.length; i++) {
          sharepointFields.push(rows[i].SharePointField);
          laserficheFields.push(rows[i].LaserficheField);
        }
      }
      this.GetItemIdByTitle().then((results: IListItem[]) => {
        this.setState({ listItem: results });
        if (this.state.listItem != null) {
          let itemId = this.state.listItem[0].Id;
          const jsonValue = JSON.parse(this.state.listItem[0].JsonValue);
          if (jsonValue.length > 0) {
            for (var i = 0; i < jsonValue.length; i++) {
              if (jsonValue[i].ConfigurationName == this.props.match.params.name) {
                jsonValue[i].DocumentName = documentName;
                jsonValue[i].DocumentTemplate = docTemp;
                jsonValue[i].DestinationPath = destPath;
                jsonValue[i].EntryId = entryId;
                jsonValue[i].Action = action;
                jsonValue[i].SharePointFields = sharepointFields;
                jsonValue[i].LaserficheFields = laserficheFields;
                break;
              }
            }
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
            this.props.context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options).then((response: SPHttpClientResponse): void => {
              this.setState(() => { return { showConfirmModal: true }; });
            });
          }
        }
      });
    }
  }
  }

  //Confirm delete button on delete modal pop up
  public ConfirmButton() {
    history.back();
    this.setState(() => { return { showConfirmModal: false }; });
  }

  //Selected document token from Token modal pop up
  public SelectDocumentToken() {
    this.setState(() => { return { showtokensModal: true }; });
  }

  //Close Tokens modal pop up
  public CloseTokenModalUp() {
    this.setState(() => { return { showtokensModal: false }; });
  }

  //Selecting document token from Token modal pop up
  public SelectedDocumentToken() {
    let tokenSelected = $("#tkn1 option:selected").text();
    var cursorPos = document.getElementById("documentName")["selectionStart"];
    let textAreaTxt = document.getElementById("documentName")["value"];
    $('#documentName').val(textAreaTxt.substring(0, cursorPos) + tokenSelected + textAreaTxt.substring(cursorPos));
    this.setState(() => { return { showtokensModal: false }; });
  }

  //Dynamically creating SharePoint and Laserfiche select elemnts
  public renderTableData() {
    let sharePointFields = this.state.sharePointFields.map(v => (
      <option value={v.InternalName}>{v.DisplayName}</option>
    ));
    let laserficheRequiredFields = this.state.laserficheFields.map((requiredItem) => {
      if (requiredItem.InternalName.includes("[Required:true]")) {
        return (<option value={requiredItem.InternalName}>{requiredItem.DisplayName}</option>);
      }
    });
    let laserficheFields = this.state.laserficheFields.map((items) => {
      if (items.InternalName.includes("[Required:false]")) {
        return (<option value={items.InternalName}>{items.DisplayName}</option>);
      }
    });
    return this.state.mappingList.map((item, index) => {
      let laserfieldValue = this.state.mappingList[index].LaserficheField;
      if (laserfieldValue.includes("[Required:true]")) {
        return (
          <tr id={index} key={index}>
            <td>
              <select name="SharePointField" className="custom-select" value={this.state.mappingList[index].SharePointField} id={this.state.mappingList[index].id} onChange={this.handleChange(index)}>
                <option>Select</option>
                {sharePointFields}
              </select>
            </td>
            <td>
              <select name="LaserficheField" disabled className="custom-select" value={this.state.mappingList[index].LaserficheField} id={this.state.mappingList[index].id} onChange={this.handleChange(index)}>
                <option>Select</option>
                {laserficheRequiredFields}
              </select>
            </td>
            <td>
              <span style={{ fontSize: "13px", color: "red" }}>Required field in Laserfiche</span>
              <span id={'a'+index}  style={{"display":"none","color":"red","fontSize":"13px","marginLeft":"10px"}} title={""}><span className="fa fa-exclamation-triangle fa-lg" style={{"marginRight":"7px"}}></span>Data types mismatch</span>
            </td>
          </tr>
        );
      }
      else {
        return (
          <tr id={index} key={index}>
            <td>
              <select name="SharePointField" className="custom-select" value={this.state.mappingList[index].SharePointField} id={this.state.mappingList[index].id} onChange={this.handleChange(index)}>
                <option>Select</option>
                {sharePointFields}
              </select>
            </td>
            <td>
              <select name="LaserficheField" className="custom-select" value={this.state.mappingList[index].LaserficheField} id={this.state.mappingList[index].id} onChange={this.handleChange(index)}>
                <option>Select</option>
                {laserficheFields}
              </select>
            </td>
            <td>
              <a href="javascript:;" className="ml-3" onClick={this.RemoveSpecificMapping(index)}><span className="fa fa-trash-alt fa-lg mt-2"></span></a>
              <span id={'a'+index} style={{"display":"none","color":"red","fontSize":"13px","marginLeft":"10px"}} title={""}><span className="fa fa-exclamation-triangle fa-lg" style={{"marginRight":"7px"}}></span>Data types mismatch</span>
            </td>
          </tr>
        );
      }
    });
  }

  //Open the folder modal dialog box
  public async OpenFoldersModal() {
    this.setState(() => { return { showFolderModal: true }; });
    await this.initializeTreeAsync();
    this.setState({
      shouldShowOpen: this.getShouldShowOpen(),
      shouldShowSelect: this.getShouldShowSelect(),
      shouldDisableSelect: this.getShouldDisableSelect()
    });
  }

  //Close the folder modal dialog box
  public CloseFolderModalUp() {
    this.setState(() => { return { showFolderModal: false }; });
  }
  
  public render(): React.ReactElement {
    let laserficheTemplate = this.state.laserficheTemplates.map(item => (
      <option value={item}>{item}</option>
    ));
    let documentName = this.state.documentNames.map(name => (
      <option value={name}>{name}</option>
    ));
    return (
      <div>
        <div style={{ display: 'none' }}>
          <lf-login redirect_uri={this.props.context.pageContext.web.absoluteUrl + this.props.laserficheRedirectPage} redirect_behavior="Replace" authorize_url_host_name={this.state.region} client_id={clientId} ref={this.loginComponent}></lf-login>
        </div>
        <div className="container-fluid p-3" style={{"maxWidth":"85%","marginLeft":"-26px"}}>
          <main className="bg-white shadow-sm">
            <div className='editPageSpinloader' hidden={this.state.loadingContent}>
              {
                !this.state.loadingContent && <Spinner size={SpinnerSize.large} label='loading' />
              }
            </div>
            <div className="p-3" hidden={this.state.hideContent}>
              <div className="card rounded-0">
                <div className="card-header d-flex justify-content-between">
                  <div>
                    <h6 className="mb-0">Profile : <span className="h5" id="configurationTitle"></span></h6>
                  </div>
                </div>
                <div className="card-body">
                  <div className="form-group row">
                    <label htmlFor="txt1" className="col-sm-2 col-form-label">Document Name</label>
                    <div className="col-sm-6">
                      <input type="text" className="form-control" id="documentName" placeholder="Document Name" disabled></input>
                    </div>
                    {/* <div className="col-sm-2" id="tokens">
                      <a href="javascript:;" className="btn btn-primary btn-sm" data-toggle="modal" data-target="#tokensModal" onClick={() => this.SelectDocumentToken()} style={{ "marginTop": "2px !important" }}>Tokens</a>
                    </div> */}
                  </div>
                  <div className="form-group row">
                    <label htmlFor="dwl2" className="col-sm-2 col-form-label">Laserfiche Template</label>
                    <div className="col-sm-6">
                      <select className="custom-select" id="documentTemplate" onChange={() => this.OnChangeTemplate()}>
                        <option>None</option>
                        {laserficheTemplate}
                      </select>
                    </div>
                  </div>
                  <div className="form-group row">
                    <label htmlFor="txt3" className="col-sm-2 col-form-label">Laserfiche Destination</label>
                    <div className="col-sm-6">
                      <input type="text" className="form-control" id="destinationPath" placeholder="(Path in Laserfiche) Example: \folder\subfolder" disabled></input>
                      <div><span>Use the Browse button to select a path</span></div>
                      <input type="text" className="form-control" id="entryId" placeholder="(Path in Laserfiche) \Added from SharePoint" style={{ display: "none" }}></input>
                    </div>
                    <div className="col-sm-2" id="folderModal" style={{ "marginTop": "2px" }}>
                      <a href="javascript:;" className="btn btn-primary btn-sm" data-toggle="modal" data-target="#folderModal" onClick={() => this.OpenFoldersModal()} >Browse</a>
                    </div>
                  </div>
                  <div className="form-group row">
                    <label htmlFor="dwl4" className="col-sm-2 col-form-label">After import</label>
                    <div className="col-sm-6">
                      <select className="custom-select" id="action">
                      <option value={"Copy"}>Leave a copy of the file in SharePoint</option>
                        <option value={"Replace"}>Replace SharePoint file with a link to the document in Laserfiche</option>
                        <option value={"Move and Delete"}>Delete SharePoint file</option>
                      </select>
                    </div>
                    <div className="col-sm-2">
                      {/* <div className="custom-control custom-checkbox mt-2" style={{ paddingLeft: "3px !important", "marginLeft": "-23px" }}>
                        <a data-toggle="tooltip" style={{ "color": "#0062cc" }}><span className="fa fa-question-circle fa-2"></span></a>
                      </div> */}
                    </div>
                  </div>
                </div>
                <h6 className="card-header border-top">Mappings from SharePoint Column to Laserfiche Field Values</h6>
                <div className="card-body">
                  <table className="table table-sm">
                    <thead>
                      <tr>
                        <th className="text-center" style={{ width: "39%" }}>SharePoint Column</th>
                        <th className="text-center" style={{ width: "38%" }}>Laserfiche Field</th>
                        <th></th>
                      </tr>
                    </thead>
                    <tbody id="tableEditBodyId">
                      {this.renderTableData()}
                    </tbody>
                  </table>
                </div>
                <div id="sharePointFieldMapping" style={{ color: "red" }}><span>Select a content type{/*  from the SharePoint column drop down instead of default "Select" value */}</span></div>
                <div id="laserficheFieldMapping" style={{ color: "red" }}><span>Select a content type{/*  from the Laserfiche field drop down instead of default "Select" value */}</span></div>
                <div id="addMapping" style={{ color: "red" }}><span>Please select any template from Laserfiche Template to add new mapping</span></div>
                <div className="card-footer bg-transparent">
                  <NavLink id="navid" to="/ManageConfigurationsPage"><a className="btn btn-primary pl-5 pr-5 float-right ml-2">Back</a></NavLink>
                  <a onClick={() => this.AddNewMappingFields()} className="btn btn-primary pl-5 pr-5 float-right ml-2">Add Field</a>
                  <a href="javascript:;" className="btn btn-primary pl-5 pr-5 float-right ml-2" onClick={() => this.SaveNewManageConfigurtaion()}>Save</a>
                </div>
              </div>
            </div>
          </main>
        </div>
        <div className="modal" id="tokensModal" hidden={!this.state.showtokensModal} data-backdrop="static" data-keyboard='false'>
          <div className="modal-dialog modal-dialog-centered">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="ModalLabel">Tokens</h5>
                <button type="button" className="close" data-dismiss="modal" aria-label="Close" onClick={() => this.CloseTokenModalUp()}>
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
              <div className="modal-body">
                <p>Select the token form the list box below</p>
                <select className="form-control" id='tkn1'>
                  {documentName}
                </select>
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-primary btn-sm" data-dismiss="modal" onClick={() => this.SelectedDocumentToken()}>Select</button>
                <button type="button" className="btn btn-secondary btn-sm" data-dismiss="modal" onClick={() => this.CloseTokenModalUp()} >Cancel</button>
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
                Do you want to delete the field mapping?
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-primary btn-sm" data-dismiss="modal" onClick={() => this.DeleteMapping()}>OK</button>
                <button type="button" className="btn btn-secondary btn-sm" data-dismiss="modal" onClick={() => this.CloseModalUp()}>Cancel</button>
              </div>
            </div>
          </div>
        </div>
        <div className="modal" id="ConfirmModal" hidden={!this.state.showConfirmModal} data-backdrop="static" data-keyboard='false'>
          <div className="modal-dialog modal-dialog-centered">
            <div className="modal-content">
              <div className="modal-body">
                Profile Updated
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-primary btn-sm" data-dismiss="modal" onClick={() => this.ConfirmButton()}>OK</button>
              </div>
            </div>
          </div>
        </div>
        <div className="modal" id="folderModal" hidden={!this.state.showFolderModal} data-backdrop="static" data-keyboard='false'>
          <div className="modal-dialog modal-dialog-centered">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="ModalLabel">Select folder for saving to Laserfiche</h5>
                <button type="button" className="close" data-dismiss="modal" aria-label="Close" onClick={() => this.CloseFolderModalUp()}>
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
              <div className="modal-body">
                <div>
                  <div ref={this.divRef}></div>
                </div>
                <div className="lf-folder-browser-sample-container" style={{ "height": "400px" }}>
                  {/* <lf-folder-browser ref={this.folderbrowser} ok_button_text="Okay" cancel_button_text="Cancel"></lf-folder-browser> */}
                  <div className="repository-browser"> 
                  <lf-repository-browser ref={this.repositoryBrowser} ok_button_text="Okay" cancel_button_text="Cancel" multiple="false"
                style={{height: '420px'}} isSelectable={this.isNodeSelectable}></lf-repository-browser>
                  <div className="repository-browser-button-containers">
                <span>
                  <button className="lf-button primary-button" onClick={this.onOpenNode} hidden={!this.state.shouldShowOpen}>OPEN
                  </button>
                  <button className="lf-button primary-button" onClick={this.onSelectFolder} hidden={!this.state.shouldShowSelect}
                  disabled={this.state.shouldDisableSelect}>Select
                  </button>
                  <button className="sec-button lf-button margin-left-button" hidden={!this.state.showFolderModal}
                  onClick={this.onClickCancelButton}>CANCEL</button>
                </span>
              </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}