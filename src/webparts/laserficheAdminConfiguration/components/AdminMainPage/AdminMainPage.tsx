import * as React from "react";
import { NavLink } from "react-router-dom";
import * as $ from "jquery";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { IAdminPageProps } from "./IAdminPageProps";
import {
  LfFieldsService
} from "@laserfiche/lf-ui-components-services";
import { LoginState } from "@laserfiche/types-lf-ui-components";
import { IRepositoryApiClientExInternal } from "../../../../repository-client/repository-client-types";
import { RepositoryClientExInternal } from "../../../../repository-client/repository-client";
import { clientId } from "../../../constants";
require("../../../../Assets/CSS/bootstrap.min.css");
require("../../../../Assets/CSS/adminConfig.css");

declare global {
  namespace JSX {
    interface IntrinsicElements {
      ["lf-login"]: any;
    }
  }
}

export default class AdminMainPage extends React.Component<IAdminPageProps, {region: string;}> {
  public loginComponent: React.RefObject<any>;
  public repoClient: IRepositoryApiClientExInternal;
  public lfFieldsService: LfFieldsService;

  constructor(props: IAdminPageProps) {
    super(props);
    SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/indigo-pink.css');
    SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ms-office-lite.css');
    this.loginComponent = React.createRef();

    this.state = {region: this.props.devMode ? 'a.clouddev.laserfiche.com' : 'laserfiche.com'};
  }
  public async componentDidMount(): Promise<void> {
    await SPComponentLoader.loadScript('https://cdn.jsdelivr.net/npm/zone.js@0.11.4/bundles/zone.umd.min.js');
    await SPComponentLoader.loadScript('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ui-components.js');
    SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/indigo-pink.css');
    SPComponentLoader.loadCss('https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@13/cdn/lf-ms-office-lite.css');

    //Add event listener to lf login component
    this.loginComponent.current.addEventListener(
      "loginCompleted",
      this.loginCompleted
    );
    this.loginComponent.current.addEventListener(
      "logoutCompleted",
      this.logoutCompleted
    );

    //Check the status of lf login state and based on that we are hiding navigation links
    const loggedOut: boolean =
      this.loginComponent.current.state === LoginState.LoggedOut;
    if (loggedOut) {
      $(".ManageConfigurationLink").hide();
      $(".ManageMappingLink").hide();
      $(".HomeLink").hide();
    } else {
      $(".ManageConfigurationLink").show();
      $(".ManageMappingLink").show();
      $(".HomeLink").show();
    }
    //await this.getAndInitializeRepositoryClientAndServicesAsync();
    //Create AdminConfiguration list in SharePoint site
    this.CreateAdminConfigList();
    //Create DocumentConfiguration list in SharePoint site
    this.CreateDocumentConfigList();
  }
  //Create a documentconfiguration list in SharePoint
  public CreateDocumentConfigList() {
    const listUrl: string =
      this.props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/GetByTitle('DocumentNameConfigList')";
    this.props.context.spHttpClient
      .get(listUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          return;
        }
        if (response.status === 404) {
          const url: string =
            this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists";
          const listDefinition: any = {
            Title: "DocumentNameConfigList",
            Description: "My description",
            BaseTemplate: 100,
          };
          const spHttpClientOptions: ISPHttpClientOptions = {
            body: JSON.stringify(listDefinition),
          };
          this.props.context.spHttpClient
            .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
            .then((responses: SPHttpClientResponse) => {
              return responses.json();
            })
            .then((responses: { value: [] }): void => {
              console.log(responses);
              var documentlist = responses["Title"];
              this.AddItemsInDocumentConfigList(documentlist);
            });
        }
      });
  }
  //Adding items in newly created document configuration list
  public AddItemsInDocumentConfigList(documentlist) {
    const arary = [
      "%(count)",
      "%(date)",
      "%(datetime)",
      "%(gcount)",
      "%(id)",
      "%(name)",
      "%(parent)",
      "%(parentid)",
      "%(parentname)",
      "%(parentuuid)",
      "%(time)",
      "%(username)",
      "%(usersid)",
      "%(uuid)",
    ];
    for (var i = 0; i < arary.length; i++) {
      let restApiUrl: string =
        this.props.context.pageContext.web.absoluteUrl +
        "/_api/web/lists/getByTitle('" +
        documentlist +
        "')/items";
      const body: string = JSON.stringify({ Title: arary[i] });
      const options: ISPHttpClientOptions = {
        headers: {
          Accept: "application/json;odata=nometadata",
          "content-type": "application/json;odata=nometadata",
          "odata-version": "",
        },
        body: body,
      };
      this.props.context.spHttpClient.post(
        restApiUrl,
        SPHttpClient.configurations.v1,
        options
      );
    }
  }
  //Creating AdminConfigurationList to store admin config values
  public CreateAdminConfigList() {
    const listUrl: string =
      this.props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/GetByTitle('AdminConfigurationList')";
    this.props.context.spHttpClient
      .get(listUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          return;
        }
        if (response.status === 404) {
          const url: string =
            this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists";
          const listDefinition: any = {
            Title: "AdminConfigurationList",
            Description: "My description",
            BaseTemplate: 100,
          };
          const spHttpClientOptions: ISPHttpClientOptions = {
            body: JSON.stringify(listDefinition),
          };
          this.props.context.spHttpClient
            .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
            .then((responses: SPHttpClientResponse) => {
              return responses.json();
            })
            .then((responses: { value: [] }): void => {
              var listtitle = responses["Title"];
              this.GetFormDigestValue(listtitle);
            });
        }
      });
  }
  //Get Form Digest Value from the context information
  public GetFormDigestValue(listtitle) {
    $.ajax({
      url: this.props.context.pageContext.web.absoluteUrl + "/_api/contextinfo",
      type: "POST",
      async: false,
      headers: { accept: "application/json;odata=verbose" },
      success: (data) => {
        var FormDigestValue = data.d.GetContextWebInformation.FormDigestValue;
        this.CreateColumns(listtitle, FormDigestValue);
      },
      error: (xhr, status, error) => {
        console.log("Failed");
      },
    });
  }
  //Create columns in the newly created admin list
  public CreateColumns(listtitle, FormDigestValue) {
    let siteUrl: string =
      this.props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getByTitle('" +
      listtitle +
      "')/fields";
    $.ajax({
      url: siteUrl,
      type: "POST",
      data: JSON.stringify({
        __metadata: { type: "SP.Field" },
        Title: "JsonValue",
        FieldTypeKind: 3,
      }),
      headers: {
        accept: "application/json;odata=verbose",
        "content-type": "application/json;odata=verbose",
        "X-RequestDigest": FormDigestValue,
      },
      success: this.onQuerySucceeded,
      error: this.onQueryFailed,
    });
  }
  public onQuerySucceeded(data) {
    console.log("Fields created");
  }
  public onQueryFailed() {
    console.log("Error!");
  }
  public loginCompleted = async () => {
    await this.getAndInitializeRepositoryClientAndServicesAsync();
    $(".ManageConfigurationLink").show();
    $(".ManageMappingLink").show();
    $(".HomeLink").show();
  }
  public logoutCompleted = async () => {
    $(".ManageConfigurationLink").hide();
    $(".ManageMappingLink").hide();
    $(".HomeLink").hide();
    window.location.href =
      this.props.context.pageContext.web.absoluteUrl +
      this.props.laserficheRedirectPage;
  }

  private async getAndInitializeRepositoryClientAndServicesAsync() {
    const accessToken =
      this.loginComponent?.current?.authorization_credentials?.accessToken;
    if (accessToken) {
      await this.ensureRepoClientInitializedAsync();
      this.lfFieldsService = new LfFieldsService(this.repoClient);

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

  public render(): React.ReactElement {
    let redirectPage =
      this.props.context.pageContext.web.absoluteUrl +
      this.props.laserficheRedirectPage;
    return (
      <div style={{ borderBottom: "3px solid #CE7A14", width: "80%" }}>
        <div className="btnSignOut">
          <lf-login
            redirect_uri={redirectPage}
            authorize_url_host_name={this.state.region}
            redirect_behavior="Replace"
            client_id={clientId}
            ref={this.loginComponent}
          ></lf-login>
        </div>
        <div>
          <span
            style={{
              marginRight: "450px",
              fontSize: "18px",
              fontWeight: "500",
            }}
          >
            Profile Editor{/* {this.props.webPartTitle} */}
          </span>
          <span className="HomeLink">
            <NavLink
              to="/HomePage"
              activeStyle={{ fontWeight: "bold", color: "red" }}
              style={{
                marginRight: "25px",
                fontWeight: "500",
                fontSize: "15px",
              }}
            >
              About
            </NavLink>
          </span>
          <span className="ManageConfigurationLink">
            <NavLink
              to="/ManageConfigurationsPage"
              activeStyle={{ fontWeight: "bold", color: "red" }}
              style={{
                marginRight: "25px",
                fontWeight: "500",
                fontSize: "15px",
              }}
            >
              Profiles
            </NavLink>
          </span>
          <span className="ManageMappingLink">
            <NavLink
              to="/ManageMappingsPage"
              activeStyle={{ fontWeight: "bold", color: "red" }}
              style={{
                marginRight: "25px",
                fontWeight: "500",
                fontSize: "15px",
              }}
            >
              Profile Mapping
            </NavLink>
          </span>
        </div>
      </div>
    );
  }
}
