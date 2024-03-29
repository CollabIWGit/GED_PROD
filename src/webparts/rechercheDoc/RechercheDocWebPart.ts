import { Version } from '@microsoft/sp-core-library';
import * as $ from 'jquery';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "bootstrap";
import 'datatables.net';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClient } from '@microsoft/sp-http';
import { sp, List, IItemAddResult, UserCustomActionScope, Items, Item, ITerm, ISiteGroup, ISiteGroupInfo } from "@pnp/sp/presets/all";
import { SPComponentLoader } from '@microsoft/sp-loader';
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';



import styles from './RechercheDocWebPart.module.scss';
import * as strings from 'RechercheDocWebPartStrings';

export interface IRechercheDocWebPartProps {
  description: string;
}

SPComponentLoader.loadCss('//cdn.datatables.net/1.10.25/css/jquery.dataTables.min.css');


export default class RechercheDocWebPart extends BaseClientSideWebPart<IRechercheDocWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private graphClient: MSGraphClient;

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      // this.user = this.context.pageContext.user;
      sp.setup({
        spfxContext: this.context,
        globalCacheDisable: true
      });

      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          this.graphClient = client;
          resolve();
        }, err => reject(err));
    });
  }

  public async render(): Promise<void> {

    if (!this.renderedOnce) {

      this.domElement.innerHTML = `
    <div class="container" style=" margin-top: 1em;">
    <div id="splistAlldocsMatching" style="box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%); padding: 1em;">

    <div id="notFound" style="display: none; text-align: center;">
    <h3>Aucun résultat ne correspond à vos critères de recherche.</h3>
     </div>

    <div id="loader" style="display: flex; align-items: center; justify-content: center; height: 100%;">
      <img src="${this.context.pageContext.web.absoluteUrl}/SiteAssets/images/logoGed.png" id="logoGedBeat" alt="Loading..." />
    </div>
  </div>
  </div>`;

      require('./../../common/css/doctabs.css');
      require('./../../common/js/jquery.min');
      require('./../../common/js/popper');
      require('./../../common/js/bootstrap.min');
      require('./../../common/js/main');
      // require('./../../common/css/common.css');
      require('./../../common/css/bugfix.css');
      require('./../../common/css/minBootstrap.css');
      require('./../../common/css/responsive.css');
      require('./../../common/css/forms.css');


      const loader = document.getElementById('loader');
      const keywords = this.getKeywords();

      try {
        await this.getDocs(keywords);
        $("#loader").css("display", "none");

        // loader.remove();
      } catch (error) {
        loader.innerHTML = `Error: ${error.message}`;
      }

    }
  }

  private getKeywords() {
    var queryParms = new URLSearchParams(document.location.search.substring(1));
    var myParm = queryParms.get("keywords");
    if (myParm) {
      return myParm.trim();
    }
  }

  public async getDocs(keywords: any) {

    var value1 = "FALSE";

    var response_title = await sp.web.lists.getByTitle("Documents1").items
      .select("Id,ParentID,FolderID,Title,revision,IsFolder,description,keywords, Created")
      .top(5000)
      .filter("substringof('" + keywords + "',Title) and IsFolder eq '" + value1 + "' ")
      .getAll();

    var response_keywords = await sp.web.lists.getByTitle("Documents1").items
      .select("Id,ParentID,FolderID,Title,revision,IsFolder,description,keywords, Created")
      .top(5000)
      .filter("substringof('" + keywords + "',keywords) and IsFolder eq '" + value1 + "' ")
      .getAll();

    var response = response_title.concat(response_keywords);

    const uniqueResponse = response.reduce((acc: any[], obj: any) => {
      if (!obj.revision || obj.revision === null) return acc;
      let existingObjIndex = acc.findIndex(o => o.Title === obj.Title);

      if (existingObjIndex === -1 || Number(obj.revision) > Number(acc[existingObjIndex].revision) ||
        new Date(obj.Created) > new Date(acc[existingObjIndex].Created)) {

        if (existingObjIndex !== -1) {
          acc.splice(existingObjIndex, 1);
        }

        acc.push(obj);
      }
      return acc;
    }, [])
      .sort((a: any, b: any) => {
        if (a.Title > b.Title) return 1;
        if (a.Title < b.Title) return -1;
        if (Number(a.revision) > Number(b.revision)) return -1;
        if (Number(a.revision) < Number(b.revision)) return 1;
        const dateA = new Date(a.Created);
        const dateB = new Date(b.Created);
        if (dateA > dateB) return -1;
        if (dateA < dateB) return 1;
        return 0;
      });


    console.log("RESPONSE", uniqueResponse);

    {
      //display so table
      // $("#table_version_doc").css("display", "block");

      const listContainerDocVersions: Element = this.domElement.querySelector('#splistAlldocsMatching');

      let html: string = `<table id='tbl_doc_versions' class='table table-striped' style="width: 100%;font-size: initial;" >`;

      html += `<thead>
    <tr>
      <th class="text-left">ID</th>
      <th class="text-left">Nom du document</th>
      <th class="text-left" >Url</th>
      <th class="text-left" >Dossier</th>
      <th class="text-left" >Description</th>
      <th class="text-left" >Keywords</th>
      <th class="text-left" >Revision</th>
    </tr>
  </thead>
  <tbody id="tbl_Documents1_versions_bdy">`;


      var response_doc_versions = null;
      var value1 = "FALSE";

      if (uniqueResponse.length > 0) {
        $("#table_version_doc").css("display", "block");

        //  await Promise.all(contract.map(async (result) => {
        // await response_doc_versions.forEach(async (element_version) => {

        await Promise.all(uniqueResponse.map(async (element_version) => {

          var pdfName = '';
          var urlFile = '';
          var url = '';
          var attachmentUrl = element_version.attachmentUrl;
          var titleFolder = '';
          var value2 = "TRUE";

          await sp.web.lists.getByTitle("Documents1")
            .items
            .getById(parseInt(element_version.Id))
            .attachmentFiles
            .select('FileName', 'ServerRelativeUrl')
            .get()
            .then(responseAttachments => {
              responseAttachments
                .forEach(attachmentItem => {
                  pdfName = attachmentItem.FileName;
                  urlFile = attachmentItem.ServerRelativeUrl;
                });

            });

          if (attachmentUrl == undefined || attachmentUrl == null || attachmentUrl == "") {
            url = urlFile;
          }
          else {
            url = attachmentUrl;

          }

          const allItemsFolder: any[] = await sp.web.lists.getByTitle('Documents1').items.select("FolderID,Title").filter("FolderID eq '" + element_version.ParentID + "' and IsFolder eq '" + value2 + "'").getAll();

          allItemsFolder.forEach((x) => {

            titleFolder = x.Title;

          });


          html += `
        <tr>
        <td class="text-left">${element_version.Id}</td>

        <td class="text-left"><a href="${this.context.pageContext.web.absoluteUrl}/SitePages/Document.aspx?document=${element_version.Title}&documentId=${element_version.FolderID}" target="_blank" data-interception="off">${element_version.Title}</a></td>

        <td class="text-left"> 
       ${url}          
        </td>

        <td class="text-left"><a href="${this.context.pageContext.web.absoluteUrl}/SitePages/documentation.aspx?folder=${element_version.ParentID}" target="_blank" data-interception="off">${titleFolder}</a></td>

        <td class="text-left"> 
        ${element_version.description}          
        </td>

        <td class="text-left"> 
        ${element_version.keywords}          
        </td>

        <td class="text-left"> 
        ${element_version.revision}          
        </td>
      
       `;

        }))
          .then(() => {

            html += `</tbody>
        </table>`;
            listContainerDocVersions.innerHTML += html;
          });


        var table = $("#tbl_doc_versions").DataTable(

          {

            order: [3, 'desc'],
            columnDefs: [
              {
                target: 0,
                visible: false,
                searchable: false
              }
              ,
              {
                target: 2,
                visible: false,
                searchable: false
              }

            ]

          }
        );

      }

      else  {
        $("#notFound").css("display", "block");
      }
    }

  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

}
