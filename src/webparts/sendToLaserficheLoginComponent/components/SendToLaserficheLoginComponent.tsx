import * as React from 'react';
import { ISendToLaserficheLoginComponentProps } from './ISendToLaserficheLoginComponentProps';
import { ISendToLaserficheLoginComponentState } from './ISendToLaserficheLoginComponentState';
import { SPComponentLoader } from '@microsoft/sp-loader';
import CustomDailog from './SendToLaserficheCustomDialog';
import { Navigation } from 'spfx-navigation';
import { CreateEntryResult, PostEntryWithEdocMetadataRequest, PutFieldValsRequest, FileParameter,FieldToUpdate,ValueToUpdate } from '@laserfiche/lf-repository-api-client';
import { LoginState } from '@laserfiche/types-lf-ui-components';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import { RepositoryClientExInternal } from '../../../repository-client/repository-client';
import { clientId } from '../../constants';

declare global {
  namespace JSX {
    interface IntrinsicElements {
      ['lf-login']: any;
    }
  }
}

const Siteredirecturl = window.localStorage.getItem("Siteurl");
const dialog: CustomDailog = new CustomDailog();
let filelink: string = "";
export default class SendToLaserficheLoginComponent extends React.Component<ISendToLaserficheLoginComponentProps, ISendToLaserficheLoginComponentState> {
  public loginComponent: React.RefObject<any>;
  public repoClient: IRepositoryApiClientExInternal;

  constructor(props: ISendToLaserficheLoginComponentProps) {
    super(props);
    SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/indigo-pink.css');
    SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ms-office-lite.css');
    SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jquery/3.1.1/jquery.min.js', { globalExportsName: 'jQuery' }).then((jquery: any): void => {
      SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/js/bootstrap.min.js',  { globalExportsName: 'jQuery' }).then((): void => {        
        });
      });
    this.loginComponent = React.createRef();
    this.loginComponent = React.createRef();

    this.state = {
      baseurl: '',
      filelink: '',
      filedata: '',
      accessToken: '',
      parentItemId: 1,
      repoId: '',
      region: this.props.devMode ? 'a.clouddev.laserfiche.com' : 'laserfiche.com'
    };
  }
  public async componentDidMount():Promise<void> {
    await SPComponentLoader.loadScript('https://cdn.jsdelivr.net/npm/zone.js@0.11.4/bundles/zone.umd.min.js');
    await SPComponentLoader.loadScript('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ui-components.js');
    SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/indigo-pink.css');
    SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ms-office-lite.css');
    this.loginComponent.current.addEventListener('loginCompleted', this.loginCompleted);
    this.loginComponent.current.addEventListener('logoutCompleted', this.logoutCompleted);

    const loggedOut: boolean = this.loginComponent.current.state === LoginState.LoggedOut;

    if (!loggedOut) {
      dialog.show();
      document.getElementById('remve').innerText = "You are signed in to Laserfiche";
      document.getElementById('remveHeading').innerText="Sign out";
      if(window.localStorage.getItem("LContType")==null||undefined){
        dialog.close();
      }
      this.getAndInitializeRepositoryClientAndServicesAsync().then((any)=>{
        this.GetFileData().then(async (results: any) => {
          this.setState({ filedata: results });
          var DocTemplate = window.localStorage.getItem("DocTemplate");
          var LContType = window.localStorage.getItem("LContType");
          if (LContType != "undefined" && LContType !== null) {
            if (DocTemplate != "None") {
              this.SendToLaserficheWithMetadata(); 
            } else {
              this.SendToLaserficheNoTemplate();
            }
          } else if (LContType !== null) {
            this.SendtoLaserficheNoMapping();
          } else {
            //document.getElementById('remvefile').style.display='none';
            dialog.close();
          }
        });
      });
    } else {

    }
   
  }

  public loginCompleted = async () => {
    dialog.show();
    document.getElementById('remve').innerText = "You are signed in to Laserfiche";
      document.getElementById('remveHeading').innerText="Sign out";
      if(window.localStorage.getItem("LContType")==null||undefined){
        dialog.close();
      }
  this.getAndInitializeRepositoryClientAndServicesAsync().then((any)=>{
    this.GetFileData().then(async (results: any) => {
      this.setState({ filedata: results });
      var DocTemplate = window.localStorage.getItem("DocTemplate");
      var LContType = window.localStorage.getItem("LContType");
      if (LContType != "undefined" && LContType !== null) {
        if (DocTemplate != "None") {
          this.SendToLaserficheWithMetadata(); 
        } else {
          this.SendToLaserficheNoTemplate();
        }
      } else if (LContType !== null) {
        this.SendtoLaserficheNoMapping();
      } else {
        //document.getElementById('remvefile').style.display='none';
        dialog.close();
        //alert('Please go back to library and select a file to upload');
      }
    });
  });
  }

  //Laserfiche LF logoutCompleted
  public logoutCompleted = async () => {
    //dialog.close();
    window.location.href = this.props.context.pageContext.web.absoluteUrl + this.props.laserficheRedirectPage;
  }
  private async getAndInitializeRepositoryClientAndServicesAsync() {
    const accessToken = this.loginComponent?.current?.authorization_credentials?.accessToken;
    if (accessToken) {

      await this.ensureRepoClientInitializedAsync();

      this.setState({ accessToken: this.loginComponent.current.authorization_credentials.accessToken });
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

  public async SendToLaserficheWithMetadata() {
    //document.getElementById('remvefile').innerText="";
    dialog.show();
    var Filenamewithext = window.localStorage.getItem("Filename");
    //var encodefilename=Filenamewithext.split('.')[0];
    
    const fileNameSplitByDot = (Filenamewithext as string).split(".");
    var fileExtensionPeriod = fileNameSplitByDot.pop();
    const Filenamewithoutext = fileNameSplitByDot.join(".");
    var Parentid = window.localStorage.getItem("Destinationfolder");
    var Filemetadata1 = window.localStorage.getItem("Filemetadata");
    var Filemetadata = JSON.parse(Filemetadata1);
    var Action = window.localStorage.getItem("Action");
    var Documentname = window.localStorage.getItem("Documentname");
    var docfilenamecheck = Documentname.includes('FileName');
    var Fileurl = window.localStorage.getItem("Fileurl");
    var Fileextension = window.localStorage.getItem("Fileextension");
    var Siteurl = window.localStorage.getItem("Siteurl");
    var SiteUrl = window.localStorage.getItem("SiteUrl");
    let fieldsAlone=Filemetadata['metadata']['fields'];
    let formattedFieldValues:
    | {
        [key: string]: FieldToUpdate;
      }
    | undefined = {};

  for (const key in fieldsAlone) {
    const value = fieldsAlone[key];
    formattedFieldValues[key] = new FieldToUpdate({
      ...value,
      values: value.values.map((val) => new ValueToUpdate(val)),
    });
  }

    const edocBlob: Blob = this.state.filedata as unknown as Blob;
    const parentEntryId = Number(Parentid);
    const fieldsAndMetadata = Filemetadata; /* JSON.parse(Filemetadata); */
    // TODO make sure this matches the correct format
    const request: PostEntryWithEdocMetadataRequest = new PostEntryWithEdocMetadataRequest({
      template: fieldsAndMetadata['template'],
      metadata: new PutFieldValsRequest({
        fields: formattedFieldValues,
      }),
    });
    const repoId = await this.repoClient.getCurrentRepoId();
    if (Documentname === "") {
      const electronicDocument: FileParameter = {
        fileName: Filenamewithext,
        data: edocBlob
      };
      const entryRequest = {
        repoId,
        parentEntryId,
        fileName: Filenamewithoutext,
        autoRename: true,
        electronicDocument,
        request,
        extension: fileExtensionPeriod
      };
      try {
        const entryCreateResult: CreateEntryResult = await this.repoClient.entriesClient.importDocument(entryRequest);
        var Entryid = entryCreateResult.operations.entryCreate.entryId;
        filelink = `https://app.${this.state.region}/laserfiche/DocView.aspx?db=${repoId}&docid=${Entryid}`;

        if (Action === 'Copy') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          window.localStorage.removeItem("LContType");
        }
        else if (Action === 'Move and Delete') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          this.DeleteFile(SiteUrl, Fileurl, Filenamewithext, Filenamewithoutext);
        }
        else if (Action === 'Replace') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          this.deletefileandreplace(SiteUrl, Fileurl, Filenamewithoutext, Filenamewithext, filelink, Siteurl);
        }
        else {
          // TODO what should happen?
        }
      }
      catch (error) {
        const conflict409 = error.operations.setFields.exceptions[0].statusCode === 409;
        if(conflict409) {
          var entryidConflict1=error.operations.entryCreate.entryId;
          filelink = `https://app.${this.state.region}/laserfiche/DocView.aspx?db=${repoId}&docid=${entryidConflict1}`;

          document.getElementById("it").innerHTML = 'Document uploaded to repository, updating metadata failed due to constraint mismatch<br/> <p style="color:red">The Laserfiche template and fields were not applied to this document</p>';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          window.localStorage.removeItem("LContType");
        }
        else {
          window.alert(`Error uploding file: ${JSON.stringify(error)}`);
          //window.localStorage.clear();
          window.localStorage.removeItem("LContType");
          dialog.close();
        }
      }
      
    } else if (docfilenamecheck === false) {
      const electronicDocument: FileParameter = {
        fileName: Documentname,
        data: edocBlob
      };
      const entryRequest = {
        repoId,
        parentEntryId,
        fileName: Documentname,
        autoRename: true,
        electronicDocument,
        request,
        extension: fileExtensionPeriod
      };
      try {
        const entryCreateResult: CreateEntryResult = await this.repoClient.entriesClient.importDocument(entryRequest);
        var Entryid3 = entryCreateResult.operations.entryCreate.entryId;
        filelink = `https://app.${this.state.region}/laserfiche/DocView.aspx?db=${repoId}&docid=${Entryid3}`;  

        if (Action === 'Copy') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          window.localStorage.removeItem("LContType");
        }
        else if (Action === 'Move and Delete') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          this.DeleteFile(SiteUrl, Fileurl, Filenamewithext, Filenamewithoutext);
        }
        else if (Action === 'Replace') {
          document.getElementById("it").innerHTML ='Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          this.deletefileandreplace(SiteUrl, Fileurl, Filenamewithoutext, Filenamewithext, filelink, Siteurl);
        }
        else {
          // TODO what should happen?
        }
      }
      catch (error) {
        if(error.operations.setFields.exceptions[0].statusCode === 409) {

          var entryidConflict2=error.operations.entryCreate.entryId;
          filelink = `https://app.${this.state.region}/laserfiche/DocView.aspx?db=${repoId}&docid=${entryidConflict2}`;

          document.getElementById("it").innerHTML = 'Document uploaded to repository, updating metadata failed due to constraint mismatch<br/> <p style="color:red">The Laserfiche template and fields were not applied to this document</p>';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          window.localStorage.removeItem("LContType");
        }
        else {
          window.alert(`Error uploding file: ${JSON.stringify(error)}`);
          window.localStorage.removeItem("LContType");
          dialog.close();
        }
      }
    } else {
      var DocnameReplacedwithfilename = Documentname.replace('FileName', Filenamewithoutext);
      
      const electronicDocument: FileParameter = {
        fileName: DocnameReplacedwithfilename+`.${fileExtensionPeriod}`,
        data: edocBlob
      };
      const entryRequest = {
        repoId,
        parentEntryId,
        fileName: DocnameReplacedwithfilename,
        autoRename: true,
        electronicDocument,
        request,
        extension: fileExtensionPeriod
      };
      try {
        const entryCreateResult: CreateEntryResult = await this.repoClient.entriesClient.importDocument(entryRequest);
        //dialog.show();
        var Entryid6 = entryCreateResult.operations.entryCreate.entryId;
        filelink = `https://app.${this.state.region}/laserfiche/DocView.aspx?db=${repoId}&docid=${Entryid6}`;

        if (Action === 'Copy') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          window.localStorage.removeItem("LContType");
        }
        else if (Action === 'Move and Delete') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          this.DeleteFile(SiteUrl, Fileurl, Filenamewithext, Filenamewithoutext);
        }
        else if (Action === 'Replace') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          this.deletefileandreplace(SiteUrl, Fileurl, Filenamewithoutext, Filenamewithext, filelink, Siteurl);
        }
        else {
          // TODO what to do here
        }
      }
      catch (error) {
        if(error.operations.setFields.exceptions[0].statusCode === 409) {

          var entryidConflict3=error.operations.entryCreate.entryId;
          filelink = `https://app.${this.state.region}/laserfiche/DocView.aspx?db=${repoId}&docid=${entryidConflict3}`;

          document.getElementById("it").innerHTML = 'Document uploaded to repository, updating metadata failed due to constraint mismatch<br/> <p style="color:red">The Laserfiche template and fields were not applied to this document</p>';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          window.localStorage.removeItem("LContType");
        }
        else {
          window.alert(`Error uploding file: ${JSON.stringify(error)}`);
          window.localStorage.removeItem("LContType");
          dialog.close();
        }

      }
    }
  }
  //
  public async SendToLaserficheNoTemplate() {
    //document.getElementById('remvefile').innerText="";
    dialog.show();
    var Filenamewithext = window.localStorage.getItem("Filename");
    //var encodefilename=Filenamewithext.split('.')[0];

    const fileNameSplitByDot = (Filenamewithext as string).split(".");
    var fileExtensionPeriod = fileNameSplitByDot.pop();
    const Filenamewithoutext = fileNameSplitByDot.join(".");
  
    var Parentid = window.localStorage.getItem("Destinationfolder");
    var Action = window.localStorage.getItem("Action");
    var Documentname = window.localStorage.getItem("Documentname");
    var docfilenamecheck = Documentname.includes('FileName');
    var Fileurl = window.localStorage.getItem("Fileurl");
    var Fileextension = window.localStorage.getItem("Fileextension");
    var Siteurl = window.localStorage.getItem("Siteurl");
    var SiteUrl = window.localStorage.getItem("SiteUrl");

    const edocBlob: Blob = this.state.filedata as unknown as Blob;
    const parentEntryId = Number(Parentid);
    const repoId = await this.repoClient.getCurrentRepoId();
    if (Documentname === "") {
      const electronicDocument: FileParameter = {
        fileName: Filenamewithext,
        data: edocBlob
      };
      const entryRequest = {
        repoId,
        parentEntryId,
        fileName: Filenamewithoutext,
        autoRename: true,
        electronicDocument,
        request: new PostEntryWithEdocMetadataRequest({}),
        extension: fileExtensionPeriod
      };

      try {
        const entryCreateResult: CreateEntryResult = await this.repoClient.entriesClient.importDocument(entryRequest);
        var Entryid6 = entryCreateResult.operations.entryCreate.entryId;
        filelink = `https://app.${this.state.region}/laserfiche/DocView.aspx?db=${repoId}&docid=${Entryid6}`;

        if (Action === 'Copy') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          window.localStorage.removeItem("LContType");
        }
        else if (Action === 'Move and Delete') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          this.DeleteFile(SiteUrl, Fileurl, Filenamewithext, Filenamewithoutext);
        }
        else if (Action === 'Replace') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          this.deletefileandreplace(SiteUrl, Fileurl, Filenamewithoutext, Filenamewithext, filelink, Siteurl);
        }
      }
      catch (error) {
        if(error.operations.setFields.exceptions[0].statusCode === 409) {

          var entryidConflict4=error.operations.entryCreate.entryId;
          filelink = `https://app.${this.state.region}/laserfiche/DocView.aspx?db=${repoId}&docid=${entryidConflict4}`;

          document.getElementById("it").innerHTML = 'Document uploaded to repository, updating metadata failed due to constraint mismatch<br/> <p style="color:red">The Laserfiche template and fields were not applied to this document</p>';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          window.localStorage.removeItem("LContType");
        }
        else {
          window.alert(`Error uploding file: ${JSON.stringify(error)}`);
          window.localStorage.removeItem("LContType");
          dialog.close();
        }
      }
    } else if (docfilenamecheck === false) {
      const electronicDocument: FileParameter = {
        fileName: Documentname,
        data: edocBlob
      };
      const entryRequest = {
        repoId,
        parentEntryId,
        fileName: Documentname,
        autoRename: true,
        electronicDocument,
        request: new PostEntryWithEdocMetadataRequest({}),
        extension: fileExtensionPeriod
      };

      try {
        
        const entryCreateResult: CreateEntryResult = await this.repoClient.entriesClient.importDocument(entryRequest);
        var Entryid9 = entryCreateResult.operations.entryCreate.entryId;
        filelink = `https://app.${this.state.region}/laserfiche/DocView.aspx?db=${repoId}&docid=${Entryid9}`;

        if (Action === 'Copy') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          window.localStorage.removeItem("LContType");
        }
        else if (Action === 'Move and Delete') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          this.DeleteFile(SiteUrl, Fileurl, Filenamewithext, Filenamewithoutext);
        }
        else if (Action === 'Replace') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          this.deletefileandreplace(SiteUrl, Fileurl, Filenamewithoutext, Filenamewithext, filelink, Siteurl);
        }
        else {
          // TODO
        }
      }
      catch (error) {
        if (error.operations.setFields.exceptions[0].statusCode === 409) {
          var entryidConflict5=error.operations.entryCreate.entryId;
          filelink = `https://app.${this.state.region}/laserfiche/DocView.aspx?db=${repoId}&docid=${entryidConflict5}`;

          document.getElementById("it").innerHTML = 'Document uploaded to repository, updating metadata failed due to constraint mismatch<br/> <p style="color:red">The Laserfiche template and fields were not applied to this document</p>';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          window.localStorage.removeItem("LContType");
        }
        else {
          window.alert(`Error uploding file: ${JSON.stringify(error)}`);
          window.localStorage.removeItem("LContType");
          dialog.close();
        }

      }
    } else {
      var DocnameReplacedwithfilename = Documentname.replace('FileName', Filenamewithoutext);

      const electronicDocument: FileParameter = {
        fileName: DocnameReplacedwithfilename+`.${fileExtensionPeriod}`,
        data: edocBlob
      };
      const entryRequest = {
        repoId,
        parentEntryId,
        fileName: DocnameReplacedwithfilename,
        autoRename: true,
        electronicDocument,
        request: new PostEntryWithEdocMetadataRequest({}),
        extension: fileExtensionPeriod
      };

      try {
        const entryCreateResult: CreateEntryResult = await this.repoClient.entriesClient.importDocument(entryRequest);
        var Entryid14 = entryCreateResult.operations.entryCreate.entryId;
        filelink = `https://app.${this.state.region}/laserfiche/DocView.aspx?db=${repoId}&docid=${Entryid14}`;

        if (Action === 'Copy') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          window.localStorage.removeItem("LContType");
        }
        else if (Action === 'Move and Delete') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          this.DeleteFile(SiteUrl, Fileurl, Filenamewithext, Filenamewithoutext);
        }
        else if (Action === 'Replace') {
          document.getElementById("it").innerHTML = 'Document uploaded';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          this.deletefileandreplace(SiteUrl, Fileurl, Filenamewithoutext, Filenamewithext, filelink, Siteurl);
        }
        else {
          // TODO what to do
        }
      }
      catch (error) {

        if(error.operations.setFields.exceptions[0].statusCode === 409) {

          var entryidConflict6=error.operations.entryCreate.entryId;
          filelink = `https://app.${this.state.region}/laserfiche/DocView.aspx?db=${repoId}&docid=${entryidConflict6}`;

          document.getElementById("it").innerHTML = 'Document uploaded to repository, updating metadata failed due to constraint mismatch<br/> <p style="color:red">The Laserfiche template and fields were not applied to this document</p>';
          document.getElementById("imgid").style.display = 'none';
          document.getElementById("divid").style.display = 'block';
          document.getElementById("divid1").onclick = this.Dc;
          document.getElementById("divid13").style.display = 'block';
          document.getElementById("divid13").onclick = this.viewfile;
          document.getElementById("divid14").onclick = this.Redirect;
          window.localStorage.removeItem("LContType");
        }
        else {

          window.alert(`Error uploding file: ${JSON.stringify(error)}`);
          window.localStorage.removeItem("LContType");
          dialog.close();
        }
      }
    }
  }
  //
  public async SendtoLaserficheNoMapping() {
    //document.getElementById('remvefile').innerText="";
    dialog.show();
    var Filenamewithext = window.localStorage.getItem("Filename");


    const fileNameSplitByDot = (Filenamewithext as string).split(".");
    var fileExtensionPeriod = fileNameSplitByDot.pop();
    const Filenamewithoutext = fileNameSplitByDot.join(".");
    var Fileextension = window.localStorage.getItem("Fileextension");

    const edocBlob: Blob = this.state.filedata as unknown as Blob;
    const parentEntryId = 1;

    try {
      const repoId = await this.repoClient.getCurrentRepoId();
      const electronicDocument: FileParameter = {
        fileName: Filenamewithext,
        data: edocBlob
      };
      const entryRequest = {
        repoId,
        parentEntryId,
        fileName: Filenamewithoutext,
        autoRename: true,
        electronicDocument,
        request: new PostEntryWithEdocMetadataRequest({}),
        extension: fileExtensionPeriod
      };

      const entryCreateResult: CreateEntryResult = await this.repoClient.entriesClient.importDocument(entryRequest);
      var Entryid14 = entryCreateResult.operations.entryCreate.entryId;
      filelink = `https://app.${this.state.region}/laserfiche/DocView.aspx?db=${repoId}&docid=${Entryid14}`;
  
      document.getElementById("it").innerHTML = 'Document uploaded';
      document.getElementById("imgid").style.display = 'none';
      document.getElementById("divid").style.display = 'block';
      document.getElementById("divid1").onclick = this.Dc;
      document.getElementById("divid13").style.display = 'block';
      document.getElementById("divid13").onclick = this.viewfile;
      document.getElementById("divid14").onclick = this.Redirect;
      window.localStorage.removeItem("LContType");
    }
    catch (error) {
      window.alert(`Error uploding file: ${JSON.stringify(error)}`);
      window.localStorage.removeItem("LContType");
      dialog.close();
    }
  }
  //
  public async GetFileData(): Promise<any> {
    //document.getElementById('remvefile').style.display='block';
    var Fileurl = window.localStorage.getItem("Fileurl");
    var SiteUrl = window.localStorage.getItem("SiteUrl");
    var Filenamewithext2 = window.localStorage.getItem("Filename");
    var encde = encodeURIComponent(Filenamewithext2);
    var fileur = Fileurl?.replace(Filenamewithext2, encde);
    var Filedataurl = SiteUrl + fileur;
    try {
      const res = await fetch(Filedataurl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
        },
      });
      const results = await res.blob();
      return results;
    }
    catch (error) {
      dialog.close();
      //document.getElementById('remvefile').style.display='none';
      console.log("error occured" + error);
    }
  }
  public Redirect() {
    var Siteurl1 = window.localStorage.getItem("SiteUrl");
    var Fileurl = window.localStorage.getItem("Fileurl");
    var Filenamewithext1 = window.localStorage.getItem("Filename");
    var fileeee = Fileurl.replace(Filenamewithext1, '');
    var path = Siteurl1 + fileeee;
    Navigation.navigate(path, true);
  }
  private Dc() {
    dialog.close();
  }
  //
  private viewfile() {
    window.open(filelink);
  }
  //
  public DeleteFile(SiteUrl, Fileurl, Filenamewithext, Filenamewithoutext) {
    var encde = encodeURIComponent(Filenamewithext);
    var fileur = Fileurl.replace(Filenamewithext, encde);
    var fileUrl1 = SiteUrl + fileur;
    $.ajax({
      url: fileUrl1,
      type: "DELETE",
      async: false,
      headers: {

        "Accept": "application/json;odata=verbose",
      },
      success: (data) => {
        window.localStorage.removeItem("LContType");
        //Perform further activity upon success, like displaying a notification
        alert('File deleted successfully');
      },
      error: (data) => {
        window.localStorage.removeItem("LContType");
        console.log("An error occurred. Please try again.");
        //Log error and perform further activity upon error/exception
      }
    });
  }
  //
  public deletefileandreplace(SiteUrl, Fileurl, Filenamewithoutext, Filenamewithext, docFilelink, Siteurl) {
    var encde = encodeURIComponent(Filenamewithext);
    var fileur = Fileurl.replace(Filenamewithext, encde);
    var fileUrl1 = SiteUrl + fileur;
    $.ajax({
      url: fileUrl1,
      type: "DELETE",
      async: false,
      headers: {

        "Accept": "application/json;odata=verbose",

      },
      success: (data) => {
        alert('File replaced with link successfully');
        this.GetFormDigestValue(SiteUrl, Fileurl, Filenamewithoutext, Filenamewithext, docFilelink, Siteurl);
        //Perform further activity upon success, like displaying a notification
      },
      error: (data) => {
        window.localStorage.removeItem("LContType");
        console.log("An error occurred. Please try again.");
        //Log error and perform further activity upon error/exception
      }
    });

  }
  //
  public GetFormDigestValue(SiteUrl, Fileurl, Filenamewithoutext, Filenamewithext, docFileLink, Siteurl) {
    $.ajax
      ({
        url: Siteurl + "/_api/contextinfo",
        type: "POST",
        async: false,
        headers: { "accept": "application/json;odata=verbose" },
        success: (data) => {
          var FormDigestValue = data.d.GetContextWebInformation.FormDigestValue;
          //console.log(FormDigestValue);
          this.postlink(SiteUrl, Fileurl, Filenamewithoutext, Filenamewithext, docFileLink, Siteurl, FormDigestValue);
        },
        error: (xhr, status, error) => {
          window.localStorage.removeItem("LContType");
          console.log("Failed");
        }
      });
  }
  //
  public postlink(SiteUrl, Fileurl, Filenamewithoutext, Filenamewithext, docFilelink, Siteurl, FormDigestValue) {
    var encde1 = encodeURIComponent(Filenamewithoutext);
    var path = Fileurl.replace(Filenamewithext, '');
    var AddLinkURL = Siteurl + `/_api/web/GetFolderByServerRelativeUrl('${path}')/Files/add(url='${encde1}.url',overwrite=true)`;

    $.ajax
      ({
        url: AddLinkURL,
        type: "POST",
        data: `[InternetShortcut]\nURL=${docFilelink}`,
        async: false,
        headers: {
          "content-type": "text/plain",
          "accept": "application/json;odata=verbose",
          "X-RequestDigest": FormDigestValue,
        },
        success: (data) => {
          window.localStorage.removeItem("LContType");
          console.log("Item Inserted..!!");
          console.log(data);
        },
        error: (data) => {
          window.localStorage.removeItem("LContType");
          console.log("API Error");
          console.log(data);
        }
      });

  }
  public render(): React.ReactElement {
    return (
      <div>
        <div style={{ "borderBottom": "3px solid #CE7A14", "marginBlockEnd": "32px" }}>
          <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAAAUVBMVEXSXyj////HYzL/+/T/+Or/9d+yaUa9ZT2yaUj/9OG7Zj3SXybRYCj/+/b///3LYS/OYCvEZDS2aEL/89jAZTnMYS3/8dO7Zzusa02+ZTn/78wyF0DsAAABnUlEQVR4nO3ci26CMABGYQcoLRS5OTf2/g86R+KSLYUm2vxcPB8RTYzxkADRajkcAAAAAAAAAADYgbJcusCvqdtLnhfeJR/a96X7vOriarNJ/cUtHeiTnI7p26TsY+XRZ190sXSfVyA6X7rP6xZdzeweREeTGDt3IBIdTeCUR3Q0wQOxLNf3CWSr0ZvcPYiWIFqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVV4zeok/379m9BL2HO1Ckymlky0jRQc3Kqoou4f6YHzdaLX56PRzak757/JjfDS0dbOK6HM6Paf8P3st6lVE/9mAwPOpNcnqokOIJppoookmmmiiiSaaaKKJ3k30OfTFdU3RXZ+lT6qq6rbO+k4VXQ9fvT2OrH30Zo+3u/5rUI17NO3QmdPImIduxoyrUze0khEm5w6uqZNIRKNi91Hl5661dH+tdow6wts5J//BaJPRwH6IT1NxbDJ6vVc+nrXJaAAAAADALn0DBosqnCStFi4AAAAASUVORK5CYII=" height={"46px"} width={"45px"} style={{"marginTop":"8px","marginLeft":"8px"}}></img>
          <span id="remveHeading" style={{ "marginLeft": "10px", "fontSize": "22px", "fontWeight": "bold" }}>Sign In</span>
        </div>
        <p id="remve" style={{ "textAlign": "center", "fontWeight": "600", "fontSize": "20px" }}>Welcome to Laserfiche</p>
        {/* <p id="remvefile" style={{ "textAlign": "center", "fontWeight": "600", "fontSize": "18px","display":"none" }}>Please wait while we prepare your file to upload....</p> */}
        <div style={{ "textAlign": "center" }}>
          <lf-login redirect_uri={this.props.context.pageContext.web.absoluteUrl + this.props.laserficheRedirectPage} authorize_url_host_name={this.state.region} redirect_behavior="Replace" client_id={clientId} sign_in_text="Sign in" ref={this.loginComponent}></lf-login>
        </div>
        <div>
          <div /* className="lf-component-container lf-right-button" */ style={{ "marginTop": "35px", "textAlign": "center" }}>
            <button style={{ "fontWeight": "600" }} onClick={this.Redirect}>Go Back</button>
          </div>
        </div>
      </div>
    );
  }
}
