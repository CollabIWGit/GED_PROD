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



import styles from './RechercheDocWebPart.module.scss';
import * as strings from 'RechercheDocWebPartStrings';

export interface IRechercheDocWebPartProps {
  description: string;
}

require('./../../common/css/doctabs.css');
SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.25/css/jquery.dataTables.min.css');


export default class RechercheDocWebPart extends BaseClientSideWebPart<IRechercheDocWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private graphClient: MSGraphClient;

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      // this.user = this.context.pageContext.user;
      sp.setup({
        spfxContext: this.context
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
    this.domElement.innerHTML = `
      <div id="splistAlldocsMatching" style="box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%); padding: 1em;">
        <div id="loader" style="display: flex; align-items: center; justify-content: center; height: 100%;">
          <img src="https://ncaircalin.sharepoint.com/sites/TestMyGed/SiteAssets/images/loader.gif" alt="Loading..." />
        </div>
      </div>
    `;

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


  private getKeywords() {
    var queryParms = new URLSearchParams(document.location.search.substring(1));
    var myParm = queryParms.get("keywords");
    if (myParm) {
      return myParm.trim();
    }
  }

  public async getDocs(keywords: any) {

    var value1 = "FALSE";

    var response_title = await sp.web.lists.getByTitle("Documents").items
      .select("Id,ParentID,FolderID,Title,revision,IsFolder,description,keywords")
      .top(5000)
      .filter("substringof('" + keywords + "',Title) and IsFolder eq '" + value1 + "' ")
      .getAll();

    var response_keywords = await sp.web.lists.getByTitle("Documents").items
      .select("Id,ParentID,FolderID,Title,revision,IsFolder,description,keywords")
      .top(5000)
      .filter("substringof('" + keywords + "',keywords) and IsFolder eq '" + value1 + "' ")
      .getAll();

    var response = response_title.concat(response_keywords);

    // Remove duplicates based on title and keep only the one with the highest revision value
    var uniqueResponse = response.reduce((acc, current) => {
      const existing = acc.find(item => item.Title === current.Title);
      if (!existing) {
        acc.push(current);
      } else if (current.revision > existing.revision) {
        acc[acc.indexOf(existing)] = current;
      }
      return acc;
    }, []);


    console.log("RESPONSE", response);

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
  <tbody id="tbl_documents_versions_bdy">`;



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

          await sp.web.lists.getByTitle("Documents")
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

          const allItemsFolder: any[] = await sp.web.lists.getByTitle('Documents').items.select("FolderID,Title").filter("FolderID eq '" + element_version.ParentID + "' and IsFolder eq '" + value2 + "'").getAll();

          allItemsFolder.forEach((x) => {

            titleFolder = x.Title;

          });





          html += `
        <tr>
        <td class="text-left">${element_version.Id}</td>

        <td class="text-left"><a href="https://ncaircalin.sharepoint.com/sites/TestMyGed/SitePages/Document.aspx?document=${element_version.Title}&documentId=${element_version.FolderID}" target="_blank" data-interception="off">${element_version.Title}</a></td>


        <td class="text-left"> 
       ${url}          
        </td>

 

        <td class="text-left"><a href="https://ncaircalin.sharepoint.com/sites/TestMyGed/SitePages/Home.aspx?folder=${element_version.ParentID}" target="_blank" data-interception="off">${titleFolder}</a></td>



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
