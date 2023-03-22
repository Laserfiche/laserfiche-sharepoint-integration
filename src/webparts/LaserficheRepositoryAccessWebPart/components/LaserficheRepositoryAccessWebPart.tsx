import * as React from "react";
import { ILaserficheRepositoryAccessWebPartProps } from "./ILaserficheRepositoryAccessWebPartProps";
import { ILaserficheRepositoryAccessWebPartState } from "./ILaserficheRepositoryAccessWebPartState";
import {
  DetailsList,
  SelectionMode,
  Selection,
  IColumn,
  CheckboxVisibility,
  ScrollablePane,
  ScrollbarVisibility,
  StickyPositionType,
  Sticky,
} from "office-ui-fabric-react";
import { IDocument } from "./ILaserficheRepositoryAccessDocument";
import { mergeStyleSets } from "office-ui-fabric-react/lib/Styling";
import * as $ from "jquery";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import SvgHtmlIcons from "../components/SVGHtmlIcons";
import { SPComponentLoader } from "@microsoft/sp-loader";
import {
  CreateEntryResult,
  Entry,
  PostEntryChildrenRequest,
  PostEntryChildrenEntryType,
  FileParameter,
  PostEntryWithEdocMetadataRequest,
  PutFieldValsRequest,
  FieldToUpdate,
  ValueToUpdate,
} from "@laserfiche/lf-repository-api-client";
import { LfRepoTreeNodeService, LfFieldsService } from "@laserfiche/lf-ui-components-services";
import { LoginState } from "@laserfiche/types-lf-ui-components";
import { IRepositoryApiClientExInternal } from "../../../repository-client/repository-client-types";
import { RepositoryClientExInternal } from "../../../repository-client/repository-client";
import { clientId } from "../../constants";
require("../../../../node_modules/bootstrap/dist/js/bootstrap.min.js");
require("../../../Assets/CSS/bootstrap.min.css");
require("../../../Assets/CSS/custom.css");

declare global {
  namespace JSX {
    interface IntrinsicElements {
      ["lf-field-container"]: any;
      ["lf-login"]: any;
    }
  }
}
const classNames = mergeStyleSets({
  fileHeader: {
    fontSize: "17px",
  },
});

export default class LaserficheRepositoryAccessWebPart extends React.Component<
  ILaserficheRepositoryAccessWebPartProps,
  ILaserficheRepositoryAccessWebPartState
> {
  public loginComponent: React.RefObject<any>;
  public fieldContainer: React.RefObject<any>;
  public repoClient: IRepositoryApiClientExInternal;
  public lfRepoTreeService: LfRepoTreeNodeService;
  public lfFieldsService: LfFieldsService;
  public showFieldContainer: boolean = false;

  constructor(props: ILaserficheRepositoryAccessWebPartProps, state: ILaserficheRepositoryAccessWebPartState) {
    super(props);
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/indigo-pink.css");
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ms-office-lite.css");
    this.loginComponent = React.createRef();
    this.fieldContainer = React.createRef();
    const selection: Selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
    });
    //Defing static columns in the grid
    const columns: IColumn[] = [
      {
        key: "column1",
        name: "Name",
        className: classNames.fileHeader,
        fieldName: "Name",
        minWidth: 90,
        maxWidth: 350,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        data: "string",
        onRender: (item: IDocument) => {
          var name = "    " + item.name;
          if (item.entryType == "Document") {
            let svgFileIcon = this.GetIconClassForDocExtension(item.extension);
            return (
              <a href="#" onDoubleClick={() => this.OpenSubfolders(item)}>
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  focusable="false"
                  style={{
                    height: "20px",
                    width: "20px",
                    color: "transparent",
                  }}
                >
                  <use xlinkHref={`#${svgFileIcon}`}></use>
                </svg>
                {name}
              </a>
            );
          } else {
            return (
              <a href="#" onDoubleClick={() => this.OpenSubfolders(item)}>
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  focusable="false"
                  style={{
                    height: "20px",
                    width: "20px",
                    color: "transparent",
                  }}
                >
                  <use xlinkHref={`#${"folder-20"}`}></use>
                </svg>
                {name}
              </a>
            );
          }
        },
        isPadded: true,
      },
      {
        key: "column2",
        name: "Creation Date",
        className: classNames.fileHeader,
        fieldName: "Creation Date",
        minWidth: 90,
        maxWidth: 160,
        onColumnClick: this._onColumnClick,
        data: "number",
        onRender: (item: IDocument) => {
          var creationDate = new Date(item.creationTime).toLocaleString([], {
            year: "numeric",
            month: "numeric",
            day: "numeric",
            hour: "2-digit",
            minute: "2-digit",
          });
          var result = creationDate.replace(/,/g, "");
          return <span>{result}</span>;
        },
        isPadded: true,
      },
      {
        key: "column3",
        name: "Last Modified Date",
        className: classNames.fileHeader,
        minWidth: 90,
        maxWidth: 160,
        isResizable: true,
        isCollapsible: true,
        onColumnClick: this._onColumnClick,
        data: "number",
        onRender: (item: IDocument) => {
          var modifiedDate = new Date(item.lastModifiedTime).toLocaleString([], {
            year: "numeric",
            month: "numeric",
            day: "numeric",
            hour: "2-digit",
            minute: "2-digit",
          });
          var result = modifiedDate.replace(/,/g, "");
          return <span>{result}</span>;
        },
      },
      {
        key: "column4",
        name: "Pages",
        className: classNames.fileHeader,
        minWidth: 90,
        maxWidth: 160,
        isResizable: true,
        isCollapsible: true,
        onColumnClick: this._onColumnClick,
        data: "number",
        onRender: (item: IDocument) => {
          if (item.pageCount != 0) {
            return <span>{item.pageCount}</span>;
          }
        },
      },
      {
        key: "column6",
        name: "Template",
        className: classNames.fileHeader,
        minWidth: 90,
        maxWidth: 160,
        isResizable: true,
        isCollapsible: true,
        onColumnClick: this._onColumnClick,
        data: "string",
        onRender: (item: IDocument) => {
          return <span>{item.templateName}</span>;
        },
      },
    ];

    this.state = {
      columns: columns,
      items: [],
      selectionDetails: "",
      selection: selection,
      checkeditemid: 0,
      checkeditemfolderornot: false,
      parentItemId: 0,
      loading: false,
      uploadProgressBar: false,
      fileUploadPercentage: 5,
      webClientUrl: "",
      showUploadModal: false,
      showCreateModal: false,
      showAlertModal: false,
      region: this.props.devMode ? "a.clouddev.laserfiche.com" : "laserfiche.com",
    };
  }
  public GetIconClassForDocExtension(extension) {
    switch (extension) {
      case "":
        return "document-20";
      case "ascx":
      case "aspx":
      case "cs":
      case "css":
      case "htm":
      case "html":
      case "js":
      case "jsproj":
      case "vbs":
      case "xml":
        return "edoc-code-20";
      case "avi":
      case "mov":
      case "mpeg":
      case "rm":
      case "wmv":
      case "mp4":
      case "webm":
      case "ogv":
      case "ogg":
        return "edoc-movie-20";
      case "bmp":
      case "gif":
      case "jpeg":
      case "jpg":
      case "png":
      case "tif":
      case "tiff":
        return "image-20";
      case "config":
        return "edoc-config-20";
      case "doc":
      case "docx":
      case "dot":
        return "edoc-wordprocessing-20";
      case "mdb":
      case "accdb":
        return "edoc-database-20";
      case "pdf":
        return "edoc-pdf-20";
      case "ppt":
      case "pptx":
        return "edoc-presentation-20";
      case "qfx":
        return "edoc-qfx-20";
      case "reg":
        return "edoc-registry-20";
      case "rtf":
      case "txt":
        return "edoc-text-20";
      case "wav":
      case "mp2":
      case "mp3":
      case "opus":
      case "oga":
        return "edoc-audio-20";
      case "wfx":
        return "edoc-wfx-20";
      case "csv":
      case "xls":
      case "xlsm":
      case "xlsx":
        return "edoc-spreadsheet-20";
      case "zip":
      case "gz":
      case "rar":
      case "7z":
        return "edoc-zip-20";
      case "msg":
      case "eml":
        return "email-20";
      case "lnk":
        return "link-20";
      case "lfb":
        return "edoc-briefcase-20";
      default:
        return "edoc-generic-20";
    }
  }
  //Opening the contents in the subfolder in webpart
  public async OpenSubfolders(item) {
    let itemId = item.id;
    this.setState({ parentItemId: itemId });
    const repoId = await this.repoClient.getCurrentRepoId();
    if (item.entryType == "Folder") {
      this.setState({
        loading: true,
      });
      this.BuildBreadcrumb(item);
      const subFolderItems: Entry[] = [];
      await this.repoClient.entriesClient.getEntryListingForEach({
        callback: async (listOfEntries) => {
          if (listOfEntries.value) {
            subFolderItems.push(...listOfEntries.value);
          }
          return true;
        },
        repoId,
        entryId: itemId,
        groupByEntryType: true,
        select: "name,parentId,creationTime,lastModifiedTime,entryType,templateName,pageCount,extension,id",
      });
      if (subFolderItems) {
        let entryResult = [];
        for (var i = 0; i < subFolderItems.length; i++) {
          entryResult.push(subFolderItems[i]);
        }
        this.setState({
          items: entryResult,
        });
      }
      this.setState({ loading: false });
    }
    if (item.entryType != "Folder") {
      // assign the first repoId for now, in production there is only one repository
      window.open(this.state.webClientUrl + "/DocView.aspx?db=" + repoId + "&docid=" + itemId);
    }
  }

  //Implementing Breadcrumb functionality on double click on folders
  public BuildBreadcrumb(item) {
    var liId = item.id + item.name;
    var liElement = liId.replace(/ /g, "");
    if ($("#LaserficheBreadcrumb").find("#" + item.id).length == 0) {
      $("#LaserficheBreadcrumb").append(
        "<li class='breadcrumb-item' id='" + item.id + "'><a href='#' id='" + liElement + "'>" + item.name + "</a></li>"
      );
      let clickEvent = document.getElementById(liElement);
      clickEvent.addEventListener("click", () => {
        this.DisplayItemsUnderBreadcrumb(item);
      });
    }
  }

  //Open the laserfiche contents when we click on previous folder in breadcrumb
  public async DisplayItemsUnderBreadcrumb(item) {
    this.setState({
      loading: true,
    });
    if (
      $("#LaserficheBreadcrumb")
        .find("#" + item.id)
        .nextAll("li").length > 0
    ) {
      $("#LaserficheBreadcrumb")
        .find("#" + item.id)
        .nextAll("li")
        .remove();
    }
    let itemId = item.id;
    this.setState({ parentItemId: itemId });
    const subFolderItems: Entry[] = [];
    const repoId = await this.repoClient.getCurrentRepoId();

    await this.repoClient.entriesClient.getEntryListingForEach({
      callback: async (listOfEntries) => {
        if (listOfEntries.value) {
          subFolderItems.push(...listOfEntries.value);
        }
        return true;
      },
      repoId,
      entryId: itemId,
      groupByEntryType: true,
      select: "name,parentId,creationTime,lastModifiedTime,entryType,templateName,pageCount,extension,id",
    });
    if (subFolderItems) {
      let entryResult = [];
      for (var i = 0; i < subFolderItems.length; i++) {
        entryResult.push(subFolderItems[i]);
      }
      this.setState({
        items: entryResult,
      });
    }
    this.setState({ loading: false });
  }

  //Get all items from the Laserfiche on load of webpart
  public async componentDidMount() {
    $("#folderValidation").hide();
    $("#folderExists").hide();
    $("#folderNameValidation").hide();
    $("#mainWebpartContent").hide();

    await SPComponentLoader.loadScript("https://cdn.jsdelivr.net/npm/zone.js@0.11.4/bundles/zone.umd.min.js");
    await SPComponentLoader.loadScript(
      "https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ui-components.js"
    );
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/indigo-pink.css");
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ms-office-lite.css");
    this.setState({
      loading: true,
      showUploadModal: false,
      showCreateModal: false,
      showAlertModal: false,
    });
    this.loginComponent.current.addEventListener("loginCompleted", this.loginCompleted);
    this.loginComponent.current.addEventListener("logoutCompleted", this.logoutCompleted);
    this.fieldContainer.current.addEventListener("templateSelectedChanged", this.onTemplateChange);
    this.fieldContainer.current.addEventListener("dialogOpened", this.onDialogOpened);

    const loggedOut: boolean = this.loginComponent.current.state === LoginState.LoggedOut;
    if (loggedOut) {
      $("#mainWebpartContent").hide();
    } else {
      $("#mainWebpartContent").show();
    }
    await this.getAndInitializeRepositoryClientAndServicesAsync();
  }

  //Get all Laserfiche Items on load of webpart
  public async GetAllLaserficheItemsOnLoad() {
    if (this.state.parentItemId == 0) {
      const allItemsResponse: Entry[] = [];

      const repoId = await this.repoClient.getCurrentRepoId();
      await this.repoClient.entriesClient.getEntryListingForEach({
        callback: async (listOfEntries) => {
          if (listOfEntries.value) {
            allItemsResponse.push(...listOfEntries.value);
          }
          return true;
        },
        repoId,
        entryId: 1,
        groupByEntryType: true,
        select: "name,parentId,creationTime,lastModifiedTime,entryType,templateName,pageCount,extension,id",
      });
      if (allItemsResponse) {
        let entryResult = [];
        for (var i = 0; i < allItemsResponse.length; i++) {
          entryResult.push(allItemsResponse[i]);
        }
        this.setState({
          items: entryResult,
        });
      }
      this.setState({
        parentItemId: 1,
      });
      this.setState({ loading: false });
    }
  }

  //Get all Laserfiche Items on click on Files in the breadcrumb
  public async GetAllLaserficheItems() {
    this.setState({
      loading: true,
    });
    if ($("#LaserficheBreadcrumb").find("#RepositoryFiles").nextAll("li").length > 0) {
      $("#LaserficheBreadcrumb").find("#RepositoryFiles").nextAll("li").remove();
    }
    const allItemsResponse: Entry[] = [];
    const repoId = await this.repoClient.getCurrentRepoId();

    await this.repoClient.entriesClient.getEntryListingForEach({
      callback: async (listOfEntries) => {
        if (listOfEntries.value) {
          allItemsResponse.push(...listOfEntries.value);
        }
        return true;
      },
      repoId,
      entryId: 1,
      groupByEntryType: true,
      select: "name,parentId,creationTime,lastModifiedTime,entryType,templateName,pageCount,extension,id",
    });
    if (allItemsResponse) {
      let entryResult = [];
      for (var i = 0; i < allItemsResponse.length; i++) {
        entryResult.push(allItemsResponse[i]);
      }
      this.setState({
        items: entryResult,
      });
    }
    this.setState({
      parentItemId: 1,
    });
    this.setState({ loading: false });
  }

  //Get Field Values on Selection on template
  public onTemplateChange = async (ev: Event) => {
    const templatechange = (ev as CustomEvent<number>).detail;
    await this.updateFieldValuesAsync();
  };
  public async updateFieldValuesAsync(): Promise<void> {
    try {
      const fieldValues = await this.lfFieldsService.getAllFieldDefinitionsAsync();
      await this.fieldContainer.current.updateFieldValuesAsync(fieldValues);
    } catch (error) {
      // TODO handle error
    }
  }

  //lf-login will trigger on click on Sign in to Laserfiche
  public loginCompleted = async () => {
    $("#mainWebpartContent").show();
    await this.getAndInitializeRepositoryClientAndServicesAsync();
  };

  //lf-login will trigger on click on Sign Out
  public logoutCompleted = async () => {
    $("#mainWebpartContent").hide();
  };
  public onDialogOpened = () => {
    $("div.adhoc-modal").css("height", "450px");
  };
  public async getAndInitializeRepositoryClientAndServicesAsync() {
    const accessToken = this.loginComponent?.current?.authorization_credentials?.accessToken;
    if (accessToken) {
      await this.ensureRepoClientInitializedAsync();
      // need to set repositoryId because multiple repositories are available in development apartment

      // need to create server session with the new repositoryID
      // Note: this will hopefully be removed and there will be no need to create server session explicitly

      this.lfFieldsService = new LfFieldsService(this.repoClient);
      await this.initializeFieldContainerAsync();
      this.setState({
        webClientUrl: this.loginComponent.current.account_endpoints.webClientUrl,
      });
      this.GetAllLaserficheItemsOnLoad();
    } else {
      // user is not logged in
    }
  }

  public async ensureRepoClientInitializedAsync(): Promise<void> {
    if (!this.repoClient) {
      const repoClientCreator = new RepositoryClientExInternal(this.loginComponent);
      this.repoClient = await repoClientCreator.createRepositoryClientAsync();
    }
  }

  public async initializeFieldContainerAsync() {
    this.showFieldContainer = true;
    await this.fieldContainer.current.initAsync(this.lfFieldsService);
  }

  //Get which file/folder is selected in the grid
  private _getSelectionDetails(): string {
    this.setState({
      checkeditemfolderornot: false,
    });
    this.setState({
      checkeditemid: 0,
    });
    if (this.state.selection.getSelection().length != 0) {
      const selectionCount = this.state.selection.getSelectedCount();

      this.setState({
        checkeditemid: (this.state.selection.getSelection()[0] as IDocument).id,
      });
      if ((this.state.selection.getSelection()[0] as IDocument).entryType == "Folder") {
        this.setState({
          checkeditemfolderornot: true,
        });
      }
      return;
    }
  }

  //Provding sorting on metadata columns
  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, items } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter((currCol) => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
        this.setState({
          announcedMessage: `${currColumn.name} is sorted ${
            currColumn.isSortedDescending ? "descending" : "ascending"
          }`,
        });
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = _copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      items: newItems,
    });
  };

  //Open New folder Modal Popup
  public OpenNewFolderModal() {
    $("#folderValidation").hide();
    $("#folderExists").hide();
    $("#folderNameValidation").hide();
    $("#folderName").val("");
    this.setState({ showCreateModal: true });
  }

  //Close New folder Modal Popup
  public CloseNewFolderModal() {
    $("#folderValidation").hide();
    $("#folderExists").hide();
    $("#folderNameValidation").hide();
    $("#folderName").val("");
    this.setState({ showCreateModal: false });
  }

  //Create New Folder in Repository
  public async CreateNewFolder(folderName) {
    if ($("#folderName").val() != "") {
      if (/[^ A-Za-z0-9]/.test(folderName)) {
        $("#folderValidation").hide();
        $("#folderExists").hide();
        $("#folderNameValidation").show();
      } else {
        $("#folderValidation").hide();
        $("#folderExists").hide();
        $("#folderNameValidation").hide();

        const repoId = await this.repoClient.getCurrentRepoId();
        const postEntryChildrenRequest: PostEntryChildrenRequest = new PostEntryChildrenRequest({
          entryType: PostEntryChildrenEntryType.Folder,
          name: folderName,
        });
        const requestParameters = {
          repoId,
          entryId: this.state.parentItemId,
          request: postEntryChildrenRequest,
        };
        try {
          let array = [];
          const newFolderEntry: Entry = await this.repoClient.entriesClient.createOrCopyEntry(requestParameters);

          array.push(newFolderEntry);
          this.setState({
            items: this.state.items.concat(array),
          });
          this.setState({ showCreateModal: false });
          $("#folderName").val("");
        } catch {
          $("#folderExists").show();
        }
      }
    } else {
      $("#folderValidation").show();
    }
  }

  //Open Import file Modal Popup
  public OpenImportFileModal() {
    this.fieldContainer.current.clearAsync();
    $("#fileValidation").hide();
    $("#fileSizeValidation").hide();
    $("#fileNameValidation").hide();
    $("#fileNameWithBacklash").hide();
    $("#importFileName").text("Choose file");
    $("#importFile").val("");
    $("#uploadFileID").val("");
    this.setState({ showUploadModal: true });
    $("#uploadModal .modal-footer").show();
    $(".progress").css("display", "none");
  }

  //Close Import File Modal Popup
  public CloseImportFileModal() {
    this.fieldContainer.current.clearAsync();
    $("#importFileName").text("Choose file");
    $("#importFile").val("");
    $("#uploadFileID").val("");
    this.setState({ showUploadModal: false });
    $("#fileValidation").hide();
    $("#fileSizeValidation").hide();
    $("#fileNameValidation").hide();
    $("#fileNameWithBacklash").hide();
    $("#uploadModal .modal-footer").show();
  }

  //Import file in Repository
  public async ImportFileToRepository() {
    var fileNamePeriod = "";
    var fileData = document.getElementById("importFile")["files"][0];
    //var fileDataSize = document.getElementById("importFile")["files"][0].size;
    var renameFileName;
    const repoId = await this.repoClient.getCurrentRepoId();
    //Checking file has been uploaded or not
    if (fileData != undefined) {
      var fileDataSize = document.getElementById("importFile")["files"][0].size;
      //Checking file size is not exceeding 100mb
      if (fileDataSize < 100000000) {
        //Checking file name is valid or not
        if (
          $("#importFileName").text() !=
          "." + document.getElementById("importFile")["value"].split("\\").pop().split(".")[1]
        ) {
          //Checking if user want to change the uploaded file name
          if (fileData.name != $("#importFileName").text()) {
            renameFileName = new File([fileData], $("#importFileName").text());
          } else {
            renameFileName = document.getElementById("importFile")["files"][0];
          }
          $("#fileValidation").hide();
          $("#fileSizeValidation").hide();
          $("#fileNameValidation").hide();
          $("#fileNameWithBacklash").hide();
          let fileContainsBacklash = renameFileName.name.includes("\\") ? "Yes" : "No";
          //Checking if filename contains backlash
          if (fileContainsBacklash === "No") {
            let fieldValidation = this.fieldContainer.current.forceValidation();
            //Checking field validation
            if (fieldValidation == true) {
              $("#uploadModal .modal-footer").hide();
              const fieldValues = this.fieldContainer.current.getFieldValues();
              let formattedFieldValues:
                | {
                    [key: string]: FieldToUpdate;
                  }
                | undefined = {};

              for (const key in fieldValues) {
                const value = fieldValues[key];
                formattedFieldValues[key] = new FieldToUpdate({
                  ...value,
                  values: value.values.map((val) => new ValueToUpdate(val)),
                });
              }

              const templateValue = this.getTemplateName();
              let templateName;
              if (templateValue != undefined) {
                templateName = templateValue;
              }
              $(".progress").css("display", "block");

              this.setState({
                uploadProgressBar: !this.state.uploadProgressBar,
              });
              this.setState({ fileUploadPercentage: 100 });
              const fieldsmetadata: PostEntryWithEdocMetadataRequest = new PostEntryWithEdocMetadataRequest({
                template: templateName,
                metadata: new PutFieldValsRequest({
                  fields: formattedFieldValues,
                }),
              });
              //const fileNameSplitByDot = (renameFileName.name as string).split(".");
              const fileNameWithExt = renameFileName.name as string;
              const fileNameSplitByDot = fileNameWithExt.split(".");
              var fileextensionperiod = fileNameSplitByDot.pop();
              const fileNameNoPeriod = fileNameSplitByDot.join(".");
              let parentEntryId = this.state.parentItemId;

              const file: FileParameter = {
                data: fileData,
                fileName: fileNameWithExt,
              };
              const requestParameters = {
                repoId,
                parentEntryId,
                electronicDocument: file,
                autoRename: true,
                fileName: fileNameNoPeriod,
                request: fieldsmetadata,
                extension: fileextensionperiod,
              };

              try {
                const entryCreateResult: CreateEntryResult = await this.repoClient.entriesClient.importDocument(
                  requestParameters
                );
                let result = entryCreateResult.documentLink;
                let parentId = parseInt(result.split("Entries/")[1]);
                let entryResult = [];
                const entry: Entry = await this.repoClient.entriesClient.getEntry({
                  repoId,
                  entryId: parentId,
                  select: "name,parentId,creationTime,lastModifiedTime,entryType,templateName,pageCount,extension,id",
                });
                entryResult.push(entry);
                this.setState({
                  items: this.state.items.concat(entryResult),
                });
                this.setState({ showUploadModal: false });
              } catch (error) {
                window.alert("Error uploding file:" + JSON.stringify(error));
              }
            } else {
              this.fieldContainer.current.forceValidation();
            }
          } else {
            $("#fileNameWithBacklash").show();
          }
        } else {
          $("#fileNameValidation").show();
        }
      } else {
        $("#fileSizeValidation").show();
      }
    } else {
      $("#fileValidation").show();
    }
  }
  public getTemplateName() {
    const templateValue = this.fieldContainer.current.getTemplateValue();
    if (templateValue) {
      return templateValue.name;
    }
    return undefined;
  }

  //Set the input file Name
  public SetImportFileName() {
    var fileNamee = "";
    var fileSize = document.getElementById("importFile")["files"][0].size;
    var filenameLength = document.getElementById("importFile")["value"].split("\\").pop().split(".").length;
    for (let j = 0; j < filenameLength - 1; j++) {
      var fileSplitValue = document.getElementById("importFile")["value"].split("\\").pop().split(".")[j];
      fileNamee += fileSplitValue + ".";
    }
    if (fileSize < 100000000) {
      $("#importFileName").text(document.getElementById("importFile")["value"].split("\\").pop());
      //$('#uploadFileID').val(document.getElementById('importFile')["value"].split('\\').pop().split(".")[0]);
      $("#uploadFileID").val(fileNamee.slice(0, -1));
      $("#fileValidation").hide();
      $("#fileSizeValidation").hide();
      $("#fileNameValidation").hide();
      $("#fileNameWithBacklash").hide();
    } else {
      $("#importFileName").text(document.getElementById("importFile")["value"].split("\\").pop());
      //$('#uploadFileID').val(document.getElementById('importFile')["value"].split('\\').pop().split(".")[0]);
      $("#uploadFileID").val(fileNamee.slice(0, -1));
      $("#fileSizeValidation").show();
    }
  }

  public SetNewFileName = () => (e) => {
    var fileNamee = "";
    var filenameLength = document.getElementById("importFile")["value"].split("\\").pop().split(".").length;
    var fileExtension = document.getElementById("importFile")["value"].split("\\").pop().split(".")[filenameLength - 1];
    for (let k = 0; k < filenameLength - 1; k++) {
      var fileSplitValue = document.getElementById("importFile")["value"].split("\\").pop().split(".")[k];
      fileNamee += fileSplitValue + ".";
    }
    let importFileName = fileNamee.slice(0, -1);
    //let importFileName = document.getElementById('importFile')["value"].split('\\').pop().split(".")[0];
    let fileChangeName = $("#uploadFileID").val();
    if (importFileName != fileChangeName) {
      $("#importFileName").text(
        fileChangeName +
          "." +
          fileExtension /* document.getElementById('importFile')["value"].split('\\').pop().split(".")[1] */
      );
    }
  };

  //On scroll display remaining items
  public ScrollToDisplayLazyLoadItems = (e) => {
    if (e.target.scrollHeight == e.target.clientHeight) {
    } else if (e.target.scrollHeight - parseInt(e.target.scrollTop) == e.target.clientHeight) {
      this.GetLazyLoadItems();
      {
        /*this.GetLaserficheLazyLoadItems(this.state.accessToken, this.props.laserficheApiUrl, itemslength, itemId, this.state.repoId).then((results: IDocument[]) => {
        this.setState({ items: this.state.items.concat(results), parentItemId: itemId });
      });*/
      }
    }
  };

  public async GetLazyLoadItems() {
    var itemslength = this.state.items.length;
    let itemId = this.state.parentItemId;
    const repoId = await this.repoClient.getCurrentRepoId();
    if (itemId === 0) {
      const allItemsResponse: Entry[] = [];

      await this.repoClient.entriesClient.getEntryListingForEach({
        callback: async (listOfEntries) => {
          if (listOfEntries.value) {
            allItemsResponse.push(...listOfEntries.value);
          }
          return true;
        },
        repoId,
        entryId: itemId,
        groupByEntryType: true,
        select: "name,parentId,creationTime,lastModifiedTime,entryType,templateName,pageCount,extension,id",
      });
      if (allItemsResponse) {
        let entryResult = [];
        for (var i = 0; i < allItemsResponse.length; i++) {
          entryResult.push(allItemsResponse[i]);
        }
        this.setState({
          items: this.state.items.concat(entryResult),
        });
        this.setState({
          parentItemId: itemId,
        });
      }
    }
  }
  //public async GetLaserficheLazyLoadItems(accessToken: string, laserficheApiUrl, itemslength, itemId, repoId): Promise<IDocument[]> {
  //let array: IDocument[] = [];
  //let restApiUrl: string = laserficheApiUrl + repoId + "/Entries/" + itemId + "/Laserfiche.Repository.Folder/children?select=name,parentId,creationTime,lastModifiedTime,entryType,templateName,pages,extension,id&$top=100&$skip=" + itemslength;
  //try {
  // const res = await fetch(restApiUrl, {
  //    method: 'GET',
  //   headers: {
  //       'Accept': 'application/json',
  //      'Content-Type': 'application/json',
  //     'Authorization': 'Bearer ' + accessToken,
  //  },
  // });
  // const results = await res.json();
  // for (var i = 0; i < results.value.length; i++) {
  //     array.push(results.value[i]);
  // }
  //return array;
  //}
  //catch (error) {
  //console.log("error occured" + error);
  //}
  //}

  //Open file button functinality to open files/folder in repository from the command bar
  public async OpenFileOrFolder(checkeditemfolderornot: boolean, checkeditemid: number) {
    const repoId = await this.repoClient.getCurrentRepoId();

    if (checkeditemfolderornot == false) {
      if (checkeditemid != 0) {
        // assign the first repoId for now, in production there is only one repository
        window.open(this.state.webClientUrl + "/DocView.aspx?db=" + repoId + "&docid=" + checkeditemid);
      } else {
        this.setState({ showAlertModal: true });
      }
    } else {
      if (checkeditemid != 0) {
        // assign the first repoId for now, in production there is only one repository
        window.open(this.state.webClientUrl + "/browse.aspx?repo=" + repoId + "#?id=" + checkeditemid);
      } else {
        this.setState({ showAlertModal: true });
      }
    }
  }

  public ConfirmAlertButton() {
    this.setState({ showAlertModal: false });
  }

  public render(): React.ReactElement<ILaserficheRepositoryAccessWebPartProps> {
    const sbBg = "#D4DBD7";
    const sbThumbBg = "#068c8e";
    return (
      <div>
        <div style={{ display: "none" }}>
          <SvgHtmlIcons />
        </div>
        <div className="container-fluid p-3" style={{ maxWidth: "100%", marginLeft: "-30px" }}>
          <div className="btnSignOut">
            <lf-login
              redirect_uri={this.props.context.pageContext.web.absoluteUrl + this.props.laserficheRedirectUrl}
              redirect_behavior="Replace"
              client_id={clientId}
              authorize_url_host_name={this.state.region}
              ref={this.loginComponent}
            ></lf-login>
          </div>
          <div>
            <main className="bg-white shadow-sm">
              <nav className="navbar navbar-dark bg-white flex-md-nowrap">
                <a className="navbar-brand pl-0" href="#">
                  <img src={require("./../../../Assets/Images/laserfiche-logo.png")} /> {this.props.webPartTitle}
                </a>
              </nav>
              <div className="p-3" id="mainWebpartContent">
                <div className="d-flex justify-content-between border p-2 file-option">
                  <div>
                    <a
                      href="javascript:;"
                      className="mr-3"
                      title="Open File"
                      onClick={() => this.OpenFileOrFolder(this.state.checkeditemfolderornot, this.state.checkeditemid)}
                    >
                      <span className="material-icons">description</span>
                    </a>
                    <span>
                      <a
                        href="javascript:;"
                        className="mr-3"
                        title="Upload File"
                        onClick={() => this.OpenImportFileModal()}
                      >
                        <span className="material-icons">upload</span>
                      </a>
                    </span>
                    <span>
                      <a
                        href="javascript:;"
                        className="mr-3"
                        title="Create Folder"
                        onClick={() => this.OpenNewFolderModal()}
                      >
                        <span className="material-icons">create_new_folder</span>
                      </a>
                    </span>
                  </div>
                  <div>
                    <div className="input-group input-group-sm" style={{ display: "none" }}>
                      <input type="text" className="form-control" placeholder="Search Filename" />
                      <div className="input-group-append">
                        <button className="btn btn-secondary " type="button" id="button-addon2">
                        <span className="material-icons">search</span>
                        </button>
                      </div>
                    </div>
                  </div>
                </div>
                <div className="parentdiv">
                  <div>
                    <nav aria-label="breadcrumb">
                      <ol
                        className="breadcrumb bg-white rounded-0 border border-top-0 border-bottom-0 mb-0 border-right-0"
                        id="LaserficheBreadcrumb"
                      >
                        <li
                          className="breadcrumb-item"
                          id="RepositoryFiles"
                          onClick={() => this.GetAllLaserficheItems()}
                        >
                          <a href="#">Files</a>
                        </li>
                      </ol>
                    </nav>
                  </div>
                  <div className="spinloader">
                    {this.state.loading && <Spinner size={SpinnerSize.medium} label="loading" labelPosition="right" />}
                  </div>
                </div>
                <div className="detailsListScrollpane">
                  <ScrollablePane
                    initialScrollPosition={0}
                    scrollbarVisibility={ScrollbarVisibility.auto}
                    onScroll={this.ScrollToDisplayLazyLoadItems}
                    styles={{
                      root: {
                        selectors: {
                          ".ms-ScrollablePane--contentContainer": {
                            scrollbarColor: `${sbThumbBg} ${sbBg}`,
                          },
                          ".ms-ScrollablePane--contentContainer::-webkit-scrollbar-track": {
                            background: sbBg,
                          },
                          ".ms-ScrollablePane--contentContainer::-webkit-scrollbar-thumb": {
                            background: sbThumbBg,
                          },
                        },
                      },
                    }}
                  >
                    <DetailsList
                      items={this.state.items}
                      columns={this.state.columns}
                      selectionMode={SelectionMode.single}
                      checkboxVisibility={CheckboxVisibility.hidden}
                      selection={this.state.selection}
                      onRenderDetailsHeader={(headerProps, defaultRender) => {
                        return (
                          <Sticky
                            stickyPosition={StickyPositionType.Header}
                            isScrollSynced={true}
                            stickyBackgroundColor="transparent"
                          >
                            {defaultRender({
                              ...headerProps,
                              styles: {
                                root: {
                                  selectors: {
                                    ".ms-DetailsHeader-cellName": {
                                      fontWeight: "bold",
                                      fontSize: 17,
                                    },
                                  },
                                  background: "#f5f5f5",
                                  borderBottom: "1px solid #ddd",
                                  paddingTop: 1,
                                },
                              },
                            })}
                          </Sticky>
                        );
                      }}
                    />
                  </ScrollablePane>
                </div>
              </div>
            </main>
          </div>
        </div>
        <div
          className="modal"
          id="uploadModal"
          data-backdrop="static"
          data-keyboard="false"
          hidden={!this.state.showUploadModal}
        >
          <div className="modal-dialog modal-dialog-scrollable modal-lg">
            <div className="modal-content" style={{ width: "724px" }}>
              <div className="modal-header">
                <h5 className="modal-title" id="ModalLabel">
                  Upload File
                </h5>
                <div
                  className="progress"
                  style={{
                    display: this.state.uploadProgressBar ? "" : "none",
                    width: "100%",
                  }}
                >
                  <div
                    className="progress-bar progress-bar-striped active"
                    style={{
                      width: this.state.fileUploadPercentage + "%",
                      backgroundColor: "orange",
                    }}
                  >
                    Uploading
                  </div>
                </div>
                {/*<button type="button" className="close" data-dismiss="modal" aria-label="Close" onClick={() => this.CloseImportFileModal()}>
                  <span aria-hidden="true">&times;</span>
                    </button>*/}
              </div>
              <div className="modal-body" style={{ height: "600px" }}>
                <div className="input-group mb-3">
                  <div className="custom-file">
                    <input
                      type="file"
                      className="custom-file-input"
                      id="importFile"
                      onChange={() => this.SetImportFileName()}
                      aria-describedby="inputGroupFileAddon04"
                      placeholder="Choose file"
                    />
                    <label className="custom-file-label" id="importFileName">
                      Choose file
                    </label>
                  </div>
                </div>
                <div id="fileValidation" style={{ color: "red" }}>
                  <span>Please select the file to upload</span>
                </div>
                <div id="fileSizeValidation" style={{ color: "red" }}>
                  <span>Please select a file below 100MB size</span>
                </div>
                <div id="fileNameValidation" style={{ color: "red" }}>
                  <span>Please provide proper name of the file</span>
                </div>
                <div id="fileNameWithBacklash" style={{ color: "red" }}>
                  <span>Please provide proper name of the file without backslash</span>
                </div>
                <div className="form-group row mb-3">
                  <label className="col-sm-2 col-form-label">Name</label>
                  <div className="col-sm-10">
                    <input type="text" className="form-control" id="uploadFileID" onChange={this.SetNewFileName()} />
                  </div>
                </div>
                <div className="card">
                  <div className="lf-component-container">
                    <lf-field-container
                      collapsible="true"
                      startCollapsed="true"
                      ref={this.fieldContainer}
                    ></lf-field-container>
                  </div>
                </div>
              </div>
              <div className="modal-footer">
                <button type="button" className="btn btn-primary btn-sm" onClick={() => this.ImportFileToRepository()}>
                  OK
                </button>
                <button type="button" className="btn btn-secondary btn-sm" onClick={() => this.CloseImportFileModal()}>
                  Cancel
                </button>
              </div>
            </div>
          </div>
        </div>
        <div
          className="modal"
          id="createModal"
          data-backdrop="static"
          data-keyboard="false"
          hidden={!this.state.showCreateModal}
        >
          <div className="modal-dialog">
            <div className="modal-content">
              <div className="modal-header">
                <h5 className="modal-title" id="ModalLabel">
                  Create Folder
                </h5>
                <button
                  type="button"
                  className="close"
                  data-dismiss="modal"
                  aria-label="Close"
                  onClick={() => this.CloseNewFolderModal()}
                >
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
              <div className="modal-body">
                <div className="form-group">
                  <label>Folder Name</label>
                  <input
                    type="text"
                    className="form-control"
                    id="folderName"
                    placeholder="Name"
                    ref={(input) => input && input.focus()}
                  />
                </div>
                <div id="folderValidation" style={{ color: "red" }}>
                  <span>Please provide folder name</span>
                </div>
                <div id="folderNameValidation" style={{ color: "red" }}>
                  <span>Invalid Name, only alphanumeric are allowed.</span>
                </div>
                <div id="folderExists" style={{ color: "red" }}>
                  <span>Object already exists</span>
                </div>
              </div>
              <div className="modal-footer">
                <button
                  type="button"
                  className="btn btn-primary btn-sm"
                  data-dismiss="modal"
                  onClick={() => this.CreateNewFolder($("#folderName").val())}
                >
                  Submit
                </button>
                <button
                  type="button"
                  className="btn btn-secondary btn-sm"
                  data-dismiss="modal"
                  onClick={() => this.CloseNewFolderModal()}
                >
                  Close
                </button>
              </div>
            </div>
          </div>
        </div>
        <div
          className="modal"
          id="AlertModal"
          data-backdrop="static"
          data-keyboard="false"
          hidden={!this.state.showAlertModal}
        >
          <div className="modal-dialog">
            <div className="modal-content">
              <div className="modal-body">Please select file/folder to open</div>
              <div className="modal-footer">
                <button
                  type="button"
                  className="btn btn-primary btn-sm"
                  data-dismiss="modal"
                  onClick={() => this.ConfirmAlertButton()}
                >
                  OK
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}
