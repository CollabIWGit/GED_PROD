import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as $ from 'jquery';

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import 'jquery-ui';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape, truncate } from '@microsoft/sp-lodash-subset';

import styles from './DocDetailsWebPart.module.scss';
import * as strings from 'DocDetailsWebPartStrings';
import { sp, List, IItemAddResult, UserCustomActionScope, Items, IItem, ISiteGroup, ISiteGroupInfo, Web, RoleDefinition, IRoleDefinition, ISiteUser, PermissionKind } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { MSGraphClient } from '@microsoft/sp-http';
import { SPComponentLoader } from '@microsoft/sp-loader';
import "bootstrap";
import 'datatables.net';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import 'jstree';
import { Navigation } from 'spfx-navigation';
import moment from 'moment';
import 'downloadjs';
import { degrees, PDFDocument, radians, rgb, rotateDegrees, rotateRadians, StandardFonts, } from 'pdf-lib/cjs/api';
import download from 'downloadjs';
import { saveAs } from 'file-saver';
import 'viewerjs/dist/viewer.css';
import { Group } from '@microsoft/microsoft-graph-types';
import { User } from '@microsoft/microsoft-graph-types';
import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable'
import ExcelJS from "exceljs";

require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');
require('./../../common/css/bugfix.css');
require('./../../common/css/doctabs.css');
require('./../../common/css/minBootstrap.css');
require('./../../common/css/responsive.css');



SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.25/css/jquery.dataTables.min.css');
// SPComponentLoader.loadScript('//cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js');
// SPComponentLoader.loadScript('https://cdn.datatables.net/buttons/2.3.6/js/dataTables.buttons.min.js');
// SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.1.3/jszip.min.js');
// SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/pdfmake.min.js');
// SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/vfs_fonts.js');
// SPComponentLoader.loadScript('https://cdn.datatables.net/buttons/2.3.6/js/buttons.html5.min.js');
// SPComponentLoader.loadScript('https://cdn.datatables.net/buttons/2.3.6/js/buttons.print.min.js');



// SPComponentLoader.loadScript("https://code.jquery.com/ui/1.12.1/jquery-ui.js");
// SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css");

//require('../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
// require('./../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');

var folders = [];
var users = [];
var groups = [];
var permission_items = [];
var users_Permission = [];
var roleDefID = [];
var parentIDArray = [];
var parentTitle = [];
var table = null;
var itemFolderId = "";

var filename_add = "";
var content_add = null;


export interface IDocDetailsWebPartProps {
  description: string;
}

declare global {
  interface Window { Viewer: any; }
}


export default class DocDetailsWebPart extends BaseClientSideWebPart<IDocDetailsWebPartProps> {

  private graphClient: MSGraphClient;
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';


  protected onInit(): Promise<void> {
    SPComponentLoader.loadScript('https://ajax.googleapis.com/ajax/libs/jquery/3.6.3/jquery.min.js');

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


  private getDocTitle() {
    var queryParms = new URLSearchParams(document.location.search.substring(1));
    var myParm = queryParms.get("document");
    if (myParm) {
      return myParm.trim();
    }
  }

  private getDocId() {
    var queryParms = new URLSearchParams(document.location.search.substring(1));
    var myParm = queryParms.get("documentId");
    if (myParm) {
      return myParm.trim();
    }
  }

  //   import { sp } from "@pnp/sp";
  // import "@pnp/sp/webs";
  // import "@pnp/sp/lists";
  // import "@pnp/sp/items";

  // import { sp, PermissionKind } from '@pnp/sp';
  // ...

  private async getGroupsByName(graphClient: MSGraphClient, displayNameStartsWith: string): Promise<Group[]> {
    const allGroups: Group[] = [];

    let nextPageUrl = `/groups?$filter=startswith(displayName,'${displayNameStartsWith}')`;
    while (nextPageUrl) {
      const response = await graphClient.api(nextPageUrl).version('v1.0').get();
      allGroups.push(...response.value);
      nextPageUrl = response["@odata.nextLink"] ?? null;
    }

    return allGroups;
  }

  private async getParentID(id: any, title: any) {

    var parentID = null;
    var folderID = null;
    var parent_title = "";
    var value2 = "FALSE";
    var value1 = "TRUE";

    let parentIDArray: any[] = [];
    let parentTitle: any[] = [];


    //var parentIDArray = [] ;

    await sp.web.lists.getByTitle('Documents1').items
      .select("ID,ParentID,FolderID")
      .filter(`FolderID eq '${id}' and Title eq '${title}' and IsFolder eq 'FALSE'`)
      .get()
      .then((results) => {

        // if (results[0].ParentID != undefined || results[0].ParentID != null) {
        parentID = results[0].ParentID;

        console.log("FIRST PARENT", parentID);
        parentIDArray.unshift(parentID);
        parentIDArray.push(parentID);
        // }

        folderID = results[0].FolderID;

        parent_title = results[0].Title;


        parentTitle.push({
          parentId_doc: folderID,
          parentTitle_doc: parent_title
        });

        console.log("Parent 1", parentID);

      });



    if (parentID == 0) {

      parentIDArray = null;
      parentTitle = null;

    }

    else {

      while (parentID != 1) {

        await sp.web.lists.getByTitle('Documents1').items.select("ID,ParentID,FolderID, Title").filter("FolderID eq '" + parentID + "' and IsFolder eq '" + value1 + "'").get().then((results) => {

          // if (results[0].ParentID != undefined || results[0].ParentID != null) {
          parentID = results[0].ParentID;
          parentIDArray.unshift(parentID);
          // }

          folderID = results[0].FolderID;
          parentIDArray.unshift(parentID);

          parent_title = results[0].Title;

          parentTitle.unshift({
            parentId_doc: folderID,
            parentTitle_doc: parent_title
          });


          console.log("Parent 2", parentID);
        });

      }


      parentIDArray.push(parseInt(this.getDocId()));



      if (parentIDArray.length > 1) {
        parentIDArray.shift();
      }

      // parentIDArray.sort(function (a, b) { return a - b });
      console.log("ArrayParent", parentIDArray);
      console.log("ArrayParent_Title", parentTitle);

    }


    // this.createPath(parentTitle);

    return { parentIDArray, parentTitle };


  }

  public async getAllUsersInGroup(groupName: any): Promise<string[]> {
    try {
      const group = await sp.web.siteGroups.getByName(groupName);
      const users = await group.users();
      const emailAddresses = users.map(user => user.Email);
      console.log(`Users in group '${groupName}': ${emailAddresses}`);
      return emailAddresses;
    } catch (error) {
      console.error(`Error getting users in group '${groupName}': ${error}`);
      return [];
    }
  }

  private async downloadDoc(fileUrl: string, fileName: any, folderID: any, filigraneText: any) {

    try {
      // Add message to the DOM to notify the user that the file is being downloaded
      const message = document.createElement('div');
      message.textContent = 'Downloading file...';
      message.style.cssText = `
        position: fixed;
        bottom: 0;
        width: 100%;
        background-color: #f54630;
        color: white;
        text-align: center;
        font-size: 24px;
        padding: 10px 0;
      `;
      document.body.appendChild(message);

      const user = await sp.web.currentUser();

      const dateDownload = Date();

      // const textWatermark = 'UNCONTROLLED COPY - Downloaded on ' + dateDownload + ' .';
      const textWatermark = filigraneText + dateDownload + ' .';


      const existingPdfBytes = await fetch(fileUrl).then(res => res.arrayBuffer());
      const pdfDoc = await PDFDocument.load(existingPdfBytes, { ignoreEncryption: true });
      console.log('pdfDoc Starting...');

      const pages = await pdfDoc.getPages();

      for (const [i, page] of Object.entries(pages)) {
        const firstPage = pages[0];

        const { width, height } = firstPage.getSize();

        const helveticaFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
        const fontSize = 16;

        page.drawText(textWatermark, {
          x: 60,
          y: 60,
          size: fontSize,
          font: helveticaFont,
          color: rgb(1, 0, 1),
          opacity: 0.4,
          rotate: degrees(55)
        });
      }

      const pdfBytes = await pdfDoc.save();

      console.log('pdfBytes: ', pdfBytes);

      // Remove the message from the DOM
      document.body.removeChild(message);

      download(pdfBytes, fileName, "application/pdf");

      await this.createAudit($("#input_number").val(), folderID, user.Title, "Telechargement");

    } catch (e) {
      alert("Cannot download this file for the following reason: " + e);

      // Remove the message from the DOM in case of an error
      const message = document.querySelector('div');
      if (message) {
        document.body.removeChild(message);
      }

      window.location.reload();
    }
  }

  private async downloadDocWithoutFili(fileUrl: string, fileName: any, folderID: any) {

    try {
      // Add message to the DOM to notify the user that the file is being downloaded
      const message = document.createElement('div');
      message.textContent = 'Downloading file...';
      message.style.cssText = `
        position: fixed;
        bottom: 0;
        width: 100%;
        background-color: #f54630;
        color: white;
        text-align: center;
        font-size: 24px;
        padding: 10px 0;
      `;
      document.body.appendChild(message);

      const user = await sp.web.currentUser();

      const dateDownload = Date();

      // const textWatermark = 'UNCONTROLLED COPY - Downloaded on ' + dateDownload + ' .';
      // const textWatermark = filigraneText + dateDownload + ' .';


      const existingPdfBytes = await fetch(fileUrl).then(res => res.arrayBuffer());
      const pdfDoc = await PDFDocument.load(existingPdfBytes, { ignoreEncryption: true });


      console.log('pdfDoc Starting...');

      const pages = await pdfDoc.getPages();

      // for (const [i, page] of Object.entries(pages)) {
      //   const firstPage = pages[0];

      //   const { width, height } = firstPage.getSize();

      //   const helveticaFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
      //   const fontSize = 16;

      //   page.drawText(textWatermark, {
      //     x: 60,
      //     y: 60,
      //     size: fontSize,
      //     font: helveticaFont,
      //     color: rgb(1, 0, 1),
      //     opacity: 0.4,
      //     rotate: degrees(55)
      //   });
      // }

      const pdfBytes = await pdfDoc.save();

      console.log('pdfBytes: ', pdfBytes);

      // Remove the message from the DOM
      document.body.removeChild(message);

      download(pdfBytes, fileName, "application/pdf");

      await this.createAudit($("#input_number").val(), folderID, user.Title, "Telechargement");

    } catch (e) {
      alert("Cannot download this file for the following reason: " + e);

      // Remove the message from the DOM in case of an error
      const message = document.querySelector('div');
      if (message) {
        document.body.removeChild(message);
      }

      window.location.reload();
    }
  }

  private async generatePdfBytes2(fileUrl: string): Promise<Uint8Array> {
    try {
      const existingPdfBytes = await fetch(fileUrl).then(res => res.arrayBuffer());
      return new Uint8Array(existingPdfBytes);
    } catch (e) {
      console.error('Failed to generate PDF bytes:', e);
      throw e;
    }
  }

  public async generateTable(groups: any, x) {
    {

      var value2 = 'FALSE';
      var title = this.getDocTitle();


      const folderInfo = await sp.web.lists.getByTitle('Documents1').items
        .select("ID,ParentID,FolderID,Title,revision,IsFolder,description,attachmentUrl,IsFiligrane,IsDownloadable")
        .top(5000)
        // .filter(`FolderID eq '${x}' and IsFolder eq '${value2}'`)
        .filter(`FolderID eq '${x}' and Title eq '${title}' and IsFolder eq 'FALSE'`)
        .getAll();

      var permission_container: Element = document.getElementById("splistDocAccessRights");

      let user_current = await sp.web.currentUser();

      console.log(user_current.Title);
      // while (permission_container.firstChild) {
      //   permission_cont ainer.removeChild(permission_container.firstChild);
      // }
      // permission_container.innerHTML = "";

      // var response = null;
      let html: string = `<table id='tbl_permission' className='table table-striped' style="width: 100%;">`;

      html += `<thead>
  
      <tr>
      <th class="text-left">Id</th>
        <th class="text-left">Nom</th>
        <th class="text-center">Droits d'accès</th>
        <th class="text-center">Actions</th>
        </tr>
        </thead>
        <tbody id="tbl_permission_bdy">
        `;


      for (const element1 of groups) {

        // if (element1.Title !== user_current.Title && element1.role !== "Full Control") {

        if (element1.role !== "Limited Access" && element1.role !== "Accès limité") {

          //  if (element1.role !== "NONE") {
          html += `
          <tr>
          <td class="text-left" id="${element1.id}">${element1.id}</td>

          <td class="text-left" id="${element1.id}_personName">${element1.title}</td>
          
          <td class="text-center" id="${element1.id}_permission_value"> ${element1.role} </td>

          <td class="text-center">
          <a id="btn${element1.id}_edit" class='buttoncss' role="button">
          
          <i class="fa-solid fa-trash"
          title="Suppression" style="
      padding-left: 0.5em;font-size: 2em;"></i>
          
          </a>
        </td>
          
          </tr>
          `;
        }

        // }
      }

      html += `</tbody>
            </table>`;

      if (permission_container.childElementCount === 0) {
        permission_container.innerHTML += html;
        var table = $("#tbl_permission").DataTable({
          columnDefs: [{
            target: 0,
            visible: false,
            searchable: false
          }]
        });


        $('#tbl_permission tbody').on('click', '.buttoncss', async (event) => {

          var data = table.row($(event.currentTarget).parents('tr')).data();

          const loader = document.createElement("div");


          if (confirm(`Voulez-vous vraiment supprimer '${data[1]}' de ce document ?\nATTENTION : l'héritage des droits pour ce fichier sera rompu.`)) {


            try {

              loader.id = "loader2";
              loader.innerHTML = "<div id='loader-spinner'></div><div id='loader-text'>Veuillez patienter... les autorisations sont en cours d'ajout.</div>"; // Add the text here
              document.body.appendChild(loader);

              var allSameDocs = await this.getAllSameDocs(folderInfo[0].Title.toString(), "FALSE");

              const list = sp.web.lists.getByTitle("Documents1");


              const _item = await list.items.getById(folderInfo[0].ID);


              await _item.breakRoleInheritance(true, false);
              await _item.roleAssignments.getById(data[0]).delete()

                .then(async () => {

                  await Promise.all(allSameDocs.map(async (item_) => {

                    if (item_.ID != folderInfo[0].ID) {

                      const _item = await list.items.getById(item_.ID);
                      await _item.breakRoleInheritance(true, false);
                      await _item.roleAssignments.getById(data[0]).delete();
                    }

                  }));

                })
                .then(async () => {

                  document.body.removeChild(loader);

                  alert("Autorisation supprimée avec succès.");


                  window.location.reload();

                });

              // try {

              //   console.log("KEY", folderInfo[0].FolderID);

              //   await sp.web.lists.getByTitle("AccessRights").items.add({
              //     Title: folderInfo[0].Title.toString(),
              //     groupName: $("#users_name").val(),
              //     permission: "NONE",
              //     FolderID: folderInfo[0].ID.toString(),
              //     PrincipleID: data[0],
              //     IsDoc: "YES"
              //     // RoleDefID: permission
              //   })
              //     .then(async () => {
              //       alert("Autorisation supprimée avec succès.");
              //       await sp.web.lists.getByTitle("Documents1").items.getById(folderInfo[0].ID).update({
              //         inheriting: "Awaiting"
              //       }).then(result => {
              //         console.log("Item updated successfully");
              //       }).catch(error => {
              //         console.log("Error updating item: ", error);
              //       });



              //     });
              // }
              // catch (e) {
              //   console.log(e.message);
              // }

            } catch (err) {
              document.body.removeChild(loader);
              alert("Erreur lors de l'application de l'autorisation au document. Veuillez réessayer.");
            }
          }

          // await this.downloadDoc(data[2], data[5], data[6], 'ARCHIVED COPY - Downloaded on ');
        });

      } else {
        // Container is not empty
      }

    }
  }

  public async checkIfUserIsRefUser(graphClient: MSGraphClient): Promise<boolean> {
    try {
      const groups = await graphClient.api('/me/transitiveMemberOf/microsoft.graph.group?$count=true&$top=999').get();
      const groupList = groups.value;
      const isRefUser = groupList.some(group => group.displayName.startsWith('MYGED_REF'));
      return isRefUser;
    } catch (error) {
      console.log(error);
      return false;
    }
  }

  public async checkIfUserIsGuestUser(graphClient: MSGraphClient): Promise<boolean> {
    try {
      const groups = await graphClient.api('/me/transitiveMemberOf/microsoft.graph.group?$count=true&$top=999').get();
      const groupList = groups.value;
      const isGuestUser = groupList.some(group => group.displayName.startsWith('MYGED_GUEST'));
      return isGuestUser;
    } catch (error) {
      console.log(error);
      return false;
    }
  }

  public async getListItemPermissions(siteUrl, listName, itemId, username, password) {
    try {

      let user_current = await sp.web.currentUser();

      const credentials = btoa(`${username}:${password}`);
      const url = `${siteUrl}/_api/Web/Lists/GetByTitle('${listName}')/Items(${itemId})/RoleAssignments?$expand=Member,RoleDefinitionBindings`;

      const response = await fetch(url, {
        headers: {
          Accept: 'application/json;odata=verbose',
          Authorization: `Basic ${credentials}`,
        },
      });

      if (!response.ok) {
        throw new Error(`Network response was not ok: ${response.status}`);
      }

      const data = await response.json();
      const permissions = [];

      for (const entry of data.d.results) {
        const userOrGroup = entry.Member;

        if (userOrGroup.PrincipalType === 1 || userOrGroup.PrincipalType === 4) {
          // User or domain group
          const principalId = userOrGroup.Id;
          const title = userOrGroup.Title;
          const roleName = entry.RoleDefinitionBindings.results[0].Name;

          // Add member to permissions
          permissions.push({ type: 'member', id: principalId, role: roleName, title: title });
        }
      }

      console.log("All Permissions on Item:", permissions);

      return { permissions };
    } catch (err) {
      console.error(err);
      return { permissions: [] };
    }
  }

  public async getBasePermissions(listId: any, docId: any): Promise<any> {
    try {
      const requestUrl = `https://ncaircalin.sharepoint.com/sites/MyGed/_api/web/lists('${listId}')/items(${docId})/effectiveBasePermissions`;
      const response = await this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);

      if (response.ok) {
        const responseJSON = await response.json();

        if (responseJSON != null && responseJSON.value != null) {
          console.log("Effective Base Permissions:", responseJSON.value);
          return responseJSON.value;
        }
      } else {
        console.error("Error retrieving base permissions. Status:", response.status);
      }
    } catch (error) {
      console.error("An error occurred:", error);
    }

    return null; // Return null in case of an error or when no permissions are found
  }

  public async getBasePermTest(siteUrl, listId, docId) {
    try {
      const url = `https://ncaircalin.sharepoint.com/sites/MyGed/_api/web/lists('${listId}')/items(${docId})/effectiveBasePermissions`;

      const response = await fetch(url, {
        headers: {
          Accept: 'application/json;odata=verbose',
        },
      });

      if (!response.ok) {
        throw new Error(`Network response was not ok: ${response.status}`);
      }

      const data = await response.json();


      //  console.log("All Site Permissions:", permissions);

      return data.value;
    } catch (err) {
      console.error(err);
      return err.message;
    }
  }

  public async getBasePermTest2(listId, docId) {
    try {
      // Configure the SharePoint context using the site URL


      // Retrieve the effective base permissions for the specific item
      const item = await sp.web.lists.getById(listId).items.getById(docId).effectiveBasePermissions.get();

      const high = item.High;
      const low = item.Low;

      return { high, low };
    } catch (err) {
      console.error(err);
      return err.message;
    }
  }

  public async render(): Promise<void> {

    if (!this.renderedOnce) {

      this.domElement.innerHTML = `

    <div class="wrapper d-flex align-items-stretch">

       <div id="loader"
            style="display: flex; align-items: center; justify-content: center; position: fixed; top: 0; left: 0; width: 100%; height: 100%; z-index: 9999; backdrop-filter: blur(5px);">
            <img src="${this.context.pageContext.web.absoluteUrl}/SiteAssets/images/logoGed.png" alt="Loading..." id="logoGedBeat" />
        </div> 

        <div id="contentDetails">
            <div class="jumbotron">
                <div class="row">
                    <div class="col-md-7 top-buffer">
                        <h2 id="h2_doc_title">
                        </h2>
                    </div>
                    <div class="text-right inline" id="view_doc">
                        <h2>
                            <a href="#" role="button" id="open_doc" title="View document"> <i class="fa-regular fa-eye"
                                    title="voir"></i></a>
    
                            <a href="#" id="delete_doc" role="button" title="delete document"> <i
                                    class="fa-solid fa-box-archive" title="Archiver le document" style="
                                padding-left: 0.5em;"></i></a>
    
                            <a href="#" id="download_doc" role="button" title="Telecharger le document"> <i
                                    class="fa-solid fa-download" title="Telecharger le document" style="
                                padding-left: 0.5em;"></i></a>
    
    
                            <label class="switch" id="switch_fav">
                                <input type="checkbox" id="bookmark-switch" style="display: none;">
                                <i class="fa-regular fa-bookmark star-icon" title="Ajouter dans marque-pages" style="padding-left: 0.5em;"></i>
                            </label>
    
    
                        </h2>
                    </div>
                </div>
            </div>
    
            <div id="doc_path">
    
            </div>
    
            <ul class="nav nav-tabs" id="myTab">
                <li class="active"><a data-toggle="tab" href="#informations">Informations</a></li>
                <li><a data-toggle="tab" href="#versions">Toutes Versions</a></li>
                <li><a data-toggle="tab" href="#access">Droits d'accès</a></li>
                <li><a data-toggle="tab" href="#notifications">Notifications</a></li>
                <li><a data-toggle="tab" href="#audit">Piste d'audit</a></li>
            </ul>
    
            <div class="tab-content">
    
                <div id="informations" class="tab-pane fade in active"
                    style=" box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%); margin: 2em; padding: 2em;">
                    <h3>Informations</h3>
    
    
                    <legend>Détails</legend>
    
    
                    <div class="row">
                        <div class="col-lg-6">
    
                            <div class="form-group">
                                <label for="input_number">Référence du document </label>
                                <input type="text" id='input_number' class='form-control' disabled>
                            </div>
    
                        </div>
    
                        <div class="col-lg-6">
                            <div class="form-group">
                                <label for="input_type_doc">Dossier</label>
                                <input type="text" class="form-control" id="input_type_doc" list='folders' disabled/>
    
                                <datalist id="folders">
                                </datalist>
                            </div>
                        </div>
    
    
                    </div>
    
                    <div class="row">
                        <div class="col-lg-6">
                            <div class="form-group">
                                <label for="input_number">Description</label>
                                <textarea id='input_description' class='form-control' rows="2"></textarea>
                            </div>
    
                        </div>
    
                        <div class="col-lg-6">
    
                            <div class="form-group">
                                <label for="input_number">Mots-clés</label>
                                <textarea id='input_keywords' class='form-control' rows="2"></textarea>
                            </div>
                        </div>
                    </div>

                    <div class="row" style="
                    margin-bottom: 2rem;
                ">
                    <div class="col-lg-6">
                    <div class="form-check">
                    <input type="checkbox" class="form-check-input" id="check1" name="option_filigrane" style="
                    width: 29px;
                    height: 19px;
                ">
                    <label class="form-check-label" for="check1">Ajouter un filigrane sur le document?</label>
                  </div>

                    </div>

                    <div class="col-lg-6">
                    <div class="form-check">
                    <input type="checkbox" class="form-check-input" id="check2" name="option_download" style="
                    width: 29px;
                    height: 19px;
                ">
                    <label class="form-check-label" for="check2">Document imprimable?</label>
                  </div>
                    </div>
                </div>
    
                    <div class="row">
    
                        <div class="col-lg-6">
                            <div class="form-group">
                                <label for="input_number">Date de diffusion</label>
                                <input type="date" id='creation_date' class='form-control'>
                            </div>
                        </div>
    
                        <div class="col-lg-6">
                            <div class="form-group" id='created_by_group'>
                                <label for="input_number">Dossier ID</label>
                                <input type="text" id='created_by' class='form-control' disabled>
                            </div>
                        </div>
    
    
                    </div>
    
                    <legend>Détails de dernière mise à jour</legend>
    
                    <div class="row">
                        <div class="col-lg-6">
                            <div class="form-group">
                                <label for="input_number">Revision</label>
                                <input type="text" id='input_revision' class='form-control'>
    
                            </div>
                        </div>
    
                        <div class="col-lg-6">
                            <div class="form-group">
                                <label for="input_number">Date de Revision</label>
                                <input id="input_reviewDate" name="myBrowser" class='form-control' type="date" disabled>
    
                            </div>
                        </div>
                    </div>
    
                    <div class="row">
                        <div class="col-lg-6">
                            <div class="form-group">
                                <label for="input_number">Fichier</label>
                                <input type="file" name="file" id="file_ammendment_update" class="form-control"
                                    style="font-size: 1em;">
    
                            </div>
                        </div>
    
                        <div class="col-lg-6">
                            <div class="form-group">
                                <label for="input_number">Chemin du fichier</label>
                                <input type="text" id='path_doc' class='form-control' disabled/>
                            </div>
                        </div>
                    </div>
    
                    <div class="row">
                        <div class="col-lg-6">
                            <div class="form-group" id='updated_by_group'>
                                <label for="input_number">Updated by :</label>
                                <input type="text" id='updated_by' class='form-control' disabled>
    
                            </div>
                        </div>
    
                        <div class="col-lg-6">
                            <div class="form-group" id='updated_time_group'>
                                <label for="input_number">Date</label>
                                <input type="text" id='updated_time' class='form-control' disabled>
    
                            </div>
                        </div>
                    </div>
    
                    <div class="line line-dashed" style="
      height: 2px;
      margin: 10px 0;
      font-size: 0;
      overflow: hidden;
      background-color: transparent;
      border-width: 0;
      border-top: 1px solid #c9cbcc;"></div>
    
                    <div class="row">
                        <div class="col-lg-8">
    
                        </div>
    
                        <div class="col-lg-4 offset-8">
                            <button type="button" class="btn btn-primary update_details_doc"
                                id='update_details_doc'>Sauvegarder</button>
                            <button type="button" class="btn btn-success" id="archive_btn" title="Archiver la revision">Archiver</button>
                            <button type="button" class="btn btn-primary" id='edit_cancel_doc'>Cancel</button>
                        </div>
                    </div>
                </div>
    
    
                <div id="versions" class="tab-pane fade">
                    <h3>Toutes Versions</h3>
    
                    <div id="splistDocVersions"
                        style="box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%); padding: 1em;">
    
                    </div>
    
                </div>
    
                <div id="access" class="tab-pane fade">
                    <h3>Droits d'accès</h3>
    
                    <div class="row" style="
      box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%);
      margin: 2em;
      padding: 2em;">
                        <div class="col-lg-12">
    
                            <div class="w3-container" id="form_access_doc">
    
                                <div class="row">
                                    <div class="col-lg-6">
    
                                        <div class="form-group">
                                            <label for="user_access_doc">Ajouter un droit d'accès utilisateur</label>
                                            <input type="text" class="form-control" id="users_name" list='users' />
    
                                            <datalist id="users">
                                                <select id="select_users"></select>
                                            </datalist>
                                        </div>
    
                                    </div>
    
                                    <div class="col-lg-4">
                                        <div class="form-group">
                                            <label for="permissions_user">Type</label>
                                            <select class='form-control' name="permissions" id="permissions_user">
                                                <!--    <option value="NONE">NONE</option> -->
                                                <option value="READ">READ</option>
                                                <option value="READ_WRITE">READ_WRITE</option>
                                                <option value="ALL">ALL</option>
                                            </select>
                                        </div>
                                    </div>
    
                                    <div class="col-lg-2" style="padding-top: 1.7em;">
                                        <div class="form-group">
                                            <button type="button" class="btn btn-primary add_user mb-2"
                                                id="add_user">Ajouter</button>
                                        </div>
                                    </div>
                                </div>
    
    
                                <div class="row">
                                    <div class="col-lg-6">
    
                                        <div class="form-group">
                                            <label for="user_access_doc">Ajouter un droit d'accès de groupe</label>
                                            <input type="text" class="form-control" id="group_name" list='group' />
    
                                            <datalist id="group">
                                                <select id="select_groups"></select>
                                            </datalist>
                                        </div>
    
                                    </div>
    
                                    <div class="col-lg-4">
                                        <div class="form-group">
                                            <label for="permissions_group">Type</label>
                                            <select class='form-control' name="permissions_group" id="permissions_group">
                                                <!--   <option value="NONE">NONE</option> -->
                                                <option value="READ">READ</option>
                                                <option value="READ_WRITE">READ_WRITE</option>
                                                <option value="ALL">ALL</option>
                                            </select>
                                        </div>
                                    </div>
    
                                    <div class="col-lg-2" style="padding-top: 1.7em;">
                                        <div class="form-group">
                                            <button type="button" class="btn btn-primary add_group mb-2"
                                                id="add_group">Ajouter</button>
                                        </div>
                                    </div>
                                </div>

                                <div class="row">
                                <div class="col-lg-6" style="padding-top: 1.7em;">
                                <div id="inherit" style="display: none;">
                                <p class="h4" id="inheritText">Ce document hérite des permissions de son parent.</p>
                                </div>
                            </div>

                                </div>

    
                                <div class="row">
                                    <div class="col-lg-3" style="padding-top: 1.7em;">
                                        <div class="form-group">
                                            <button type="button" class="btn btn-primary add_group mb-2"
                                                id="inherit_parent">Hériter les droits d'accès</button>
                                        </div>
                                    </div>
                              
                                </div>

                                <div class="row">
                                <div class="col-lg-3" style="padding-top: 1.7em;">
                                    <div class="form-group">
                                        <button type="button" class="btn btn-primary add_group mb-2"
                                            id="not_inherit_parent">Rompre les droits d'accès du parent</button>
                                    </div>
                                </div>
                          
                            </div>


    
                            </div>
    
    
                        </div>
                    </div>
    
                    <div id="splistDocAccessRights"
                        style="box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%); padding: 1em;">
    
                    </div>
                </div>
    
                <div id="notifications" class="tab-pane fade">
                    <h3>Notifications</h3>
    
                    <div class="row" style="
        box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%);
        margin: 2em;
        padding: 2em;">
                        <div class="col-lg-12">
    
                            <div class="w3-container" id="form_notif_doc">
    
                                <div class="row">
                                    <div class="col-lg-6">
    
                                        <div class="form-group">
                                            <label for="users_name_notif">Ajouter une notification utilisateur :</label>
                                            <input type="text" class="form-control" id="users_name_notif" list='users' />
    
                                            <datalist id="users">
                                                <select id="select_users"></select>
                                            </datalist>
                                        </div>
    
                                    </div>
    
    
    
                                    <div class="col-lg-3" style="padding-top: 1.7em;">
                                        <div class="form-group">
                                            <button type="button" class="btn btn-primary add_notif_user mb-2"
                                                id="add_user_notif">Ajouter</button>
                                        </div>
                                    </div>
                                </div>
    
    
                                <div class="row">
                                    <div class="col-lg-6">
    
                                        <div class="form-group">
                                            <label for="group_name_notif">Ajouter une notification de groupe :</label>
                                            <input type="text" class="form-control" id="group_name_notif"
                                                list='groups_notif' />
    
                                            <datalist id="groups_notif">
                                            </datalist>
                                        </div>
    
                                    </div>
    
    
                                    <div class="col-lg-3" style="padding-top: 1.7em;">
                                        <div class="form-group">
                                            <button type="button" class="btn btn-primary add_notif_group mb-2"
                                                id="add_group_notif">Ajouter</button>
                                        </div>
                                    </div>
                                </div>
    
                            </div>
    
    
    
    
                        </div>
                    </div>
    
                    <div id="splistDocNotifications"
                        style="box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%); padding: 1em;">
    
                    </div>
    
    
                </div>
    
                <div id="audit" class="tab-pane fade">
                    <h3>Piste d'audit</h3>
    
                    <div id="splistDocAudit"
                        style="box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%); padding: 1em;">
    
                        <button type="button" class="btn btn-success" id="btn_pdf" style="
                        margin: 1rem;
                        margin-bottom: 2em;
                    ">Pdf</button>
                        <button type="button" class="btn btn-success" id="btn_excel" style="
                        margin: 1rem;
                        margin-bottom: 2em;
                    ">Excel</button>
    
                    </div>
    
    
                </div>
    
            </div>
    
        </div>
    
    </div>
    `;

      Promise.all([
        SPComponentLoader.loadScript('https://ajax.googleapis.com/ajax/libs/jquery/3.6.3/jquery.min.js'), // jQuery
        SPComponentLoader.loadScript('//cdn.jsdelivr.net/npm/popper.js@1.14.7/dist/umd/popper.min.js'), // Popper.js
        SPComponentLoader.loadScript("https://code.jquery.com/ui/1.12.1/jquery-ui.js"), // jQuery UI
        SPComponentLoader.loadScript('//cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.4.1/js/bootstrap.min.js'), // Bootstrap
        // SPComponentLoader.loadScript('//cdn.datatables.net/1.10.25/js/jquery.dataTables.min.js'), // DataTables
      ]).then(function () {
        console.log("Scripts loaded successfully");
        // All scripts have been loaded successfully
      }).catch(function (error) {
        // An error occurred while loading the scripts
        console.error("Error loading scripts: " + error);
      });

      this.getUserGroups(this.graphClient);
      // this.requireLibraries();
      const starIcon = document.querySelector(".star-icon") as HTMLElement;

      starIcon.addEventListener("click", function () {
        if (this.classList.contains("fa-solid")) {
          this.classList.remove("fa-solid");
          this.classList.add("fa-regular");
          console.log("Star is grey");
        } else {
          this.classList.remove("fa-regular");
          this.classList.add("fa-solid");
          console.log("Star is gold");
        }
      });

      var title = this.getDocTitle();
      var docId = this.getDocId();

      var principleIdOfGroup = null;
      var principleIdOfGroup_notif = null;

      //require('./DocDetailsWebPartJS');

      const loader = document.getElementById('loader');

      this.getParentID(this.getDocId(), this.getDocTitle());

      var items = await sp.web.lists.getByTitle("Documents1").items
        .select("ID, Title, ParentID, HasUniqueRoleAssignments")
        .filter(`FolderID eq '${docId}' and Title eq '${title}' and IsFolder eq 'FALSE'`)
        .get();

      // const basePermissions = await this.getCurrentUserPermissionsForItem(items[0].ID, 'be040ad0-c2e1-45bf-9a2d-3503043fb77b');
      // console.log("High", basePermissions.High);
      // console.log("Low", basePermissions.Low);
      // const x = await this.getBasePermissions('be040ad0-c2e1-45bf-9a2d-3503043fb77b', items[0].ID);

      await this.getBasePermTest2('be040ad0-c2e1-45bf-9a2d-3503043fb77b', items[0].ID)
        .then(async result => {
          // Handle the result
          console.log('High Value:', result.high);
          console.log('Low Value:', result.low);

          const high = result.high;
          const low = result.low;

          //        full control -> high : 2147483647
          //                 low : 4294967295

          // edit -> high: 432
          //         low : 1011030767

          // read --> high: 176
          //          low: 138612833

          let user_current = await sp.web.currentUser();

          if ((high == 2147483647 && low == 4294967295) || (high == 2147483647 && low == 4294705151)) { //full control
            console.log("You have full control!");
            $("#creation_date").prop("disabled", false);
            $("#input_number").prop("disabled", false);
            $("#input_type_doc").prop("disabled", false);
            $("#input_reviewDate").prop("disabled", false);
            // $("#creation_date").prop("disabled", false);

            const { permissions } = await this.getListItemPermissions(this.context.pageContext.web.absoluteUrl, "Documents1", items[0].ID, "mgolapkhan.ext@aircalin.nc", "musharaf2897");

            const excludedItems = permissions.filter(item => item.title !== user_current.Title);

            if (excludedItems.length === 0) {

              await this.generateTable([{ type: 'member', id: '', role: 'Full Control', title: 'MYGED_ADMIN' }], Number(docId));

            } else {
              await this.generateTable(excludedItems, Number(docId));
            }

            console.log("PERMISSIONS ON ITEM", permissions);

            if (items.length > 0 && items[0] && items[0].HasUniqueRoleAssignments === false) {
              $("#inherit").css("display", "block");
              $("#inherit_parent").css("display", "none");
              $("#splistDocAccessRights").css("display", "block");
              $("#not_inherit_parent").css("display", "block");
            }

            else {
              $("#splistDocAccessRights").css("display", "block");
              $("#inherit").css("display", "none");
              $("#inherit_parent").css("display", "block");
              $("#not_inherit_parent").css("display", "none");
            }


            try {
              await this.getSiteGroups(),
                await this.getSiteGroups_notif(),
                await this.getSiteUsers();
            }
            catch (err) {
              console.log(err.message);
            }

          }
          else if (high == 432 && low == 1011030767) { //edit



            $("#access").css('display', "none");
            $("#creation_date").prop("disabled", false);
            $("#input_reviewDate").prop("disabled", false);
            //$("#creation_date").prop("disabled", false);

            await this.getSiteUsers();
            await this.getSiteGroups_notif()
          }
          else if (high == 176 && low == 138612833) { //read
            $("#update_details_doc, #edit_cancel_doc, #access, #notifications, #audit, #delete_doc, #archive_btn").css("display", "none");
            $("#input_description, #input_keywords, #input_revision, #file_ammendment_update, #check1, #check2, #creation_date, input_reviewDate").prop('disabled', true);
          }

          else {

          }
        })
        .catch(error => {
          // Handle any errors
          console.error('Error:', error);
        });

      try {

        await this._getDocDetails(parseInt(docId), title),
          // await this.checkPermission(),
          await this._getAllVersions(title, docId),
          // await this._getAllAccess(docId),
          await this._getAllAudit(docId, title),
          await this._getAllNotifications(docId),
          await this.load_folders(),
          this.fileUpload();

        $("#loader").css("display", "none");
        // $(".spinner").css("display", "none");

      } catch (error) {
        // $("#loader").html(`Error: ${error.message}`);
        $("#loader").css("display", "none");
        // $(".spinner").css("display", "none");
        console.log(error.message);
      }


      $("#input_type_doc").bind('input', () => {
        const shownVal = (document.getElementById("input_type_doc") as HTMLInputElement).value;
        // var shownVal = document.getElementById("name").value;

        const value2send = (document.querySelector<HTMLSelectElement>(`#folders option[value='${shownVal}']`) as HTMLSelectElement).dataset.value;

        const selectedOption = $(`#folders option[value='${shownVal}']`);
        const fileDir = selectedOption.data("filedirref");

        console.log(value2send);
        $("#created_by").val(value2send);
        $("#path_doc").val(fileDir);

      });


      $("#group_name").bind('input', () => {
        const shownVal = (document.getElementById("group_name") as HTMLInputElement).value;
        // var shownVal = document.getElementById("name").value;

        const value2send = (document.querySelector<HTMLSelectElement>(`#group option[value='${shownVal}']`) as HTMLSelectElement).dataset.value;
        principleIdOfGroup = value2send;
        console.log(value2send);
        //  $("#created_by").val(value2send);
      });

      $("#group_name_notif").bind('input', () => {
        const shownVal = (document.getElementById("group_name_notif") as HTMLInputElement).value;
        // var shownVal = document.getElementById("name").value;

        const value2send = (document.querySelector<HTMLSelectElement>(`#groups_notif option[value='${shownVal}']`) as HTMLSelectElement).dataset.value;
        principleIdOfGroup_notif = value2send;
        console.log(value2send);
        //  $("#created_by").val(value2send);
      });


      const inputRevision = document.getElementById("input_revision") as HTMLInputElement;

      inputRevision.addEventListener('blur', function () {
        if (this.value) {
          alert("Téléchargez votre fichier avant de continuer.");
          $('#file_ammendment_update').focus();
        }
      });


      //update document
      //add_permission user
      // $("#add_user").click(async (e) => {

      //   if ($("#users_name").val() === "") {
      //     alert("Please select a user.");
      //   }
      //   else {
      //     await this.add_permission($("#users_name").val().toString(), items[0].ID);

      //   }

      // });

      $("#add_user").click(async (e) => {
        // Disable the button


        if ($("#users_name").val() === "") {
          alert("Please select a user.");
        } else {


          $("#add_user").prop("disabled", true);

          // Show loading text or spinner
          $("#add_user").text("Loading...");

          const loader = document.createElement("div");
          loader.id = "loader2";
          loader.innerHTML = "<div id='loader-spinner'></div><div id='loader-text'>Veuillez patienter... les autorisations sont en cours d'ajout.</div>"; // Add the text here
          document.body.appendChild(loader);


          try {

            const allSameDocuments = await this.getAllSameDocs(items[0].Title, 'FALSE');

            const user: any = await sp.web.siteUsers.getByEmail($("#users_name").val().toString())();

            var selected_permission = $("#permissions_user option:selected").val();

            var permission = 0;

            if (selected_permission === "ALL") {

              permission = 1073741829;
            }

            else if (selected_permission === "READ") {
              permission = 1073741826;

            }
            else if (selected_permission === "READ_WRITE") {
              permission = 1073741830;
            }

            const list = sp.web.lists.getByTitle("Documents1");

            const _item = await list.items.getById(items[0].ID);
            await _item.breakRoleInheritance(true, false);

            const roleAssignments = _item.roleAssignments;
            await roleAssignments.add(user.Id, permission)

              .then(async () => {

                await Promise.all(allSameDocuments.map(async (item_) => {

                  const _item = await list.items.getById(item_.ID);
                  await _item.breakRoleInheritance(true, false);

                  const roleAssignments = _item.roleAssignments;
                  await roleAssignments.add(user.Id, permission);

                }));

              });

            document.body.removeChild(loader);


            //  await this.add_permission($("#users_name").val().toString(), items[0].ID);


            // console.log("ALL SAME", allDocuments)

            // for (const document of allDocuments) {

            //   if(document.ID != items[0].ID){
            //     await this.add_permission($("#users_name").val().toString(), document.ID);
            //   }
            // }

            // .then(() => {
            alert("Autorisation ajoutée à ce document avec succès.");
            // })
            // .then(() => {
            window.location.reload();
            // });
            // Reset the button state
            // $("#add_user").prop("disabled", false);
            // $("#add_user").text("Ajouter");
          } catch (err) {
            // alert(err.message);
            document.body.removeChild(loader);

            alert("Erreur lors de l'application de l'autorisation au dossier. Veuillez réessayer.");

            // Enable the button and reset text on error
            $("#add_user").prop("disabled", false);
            $("#add_user").text("Ajouter");
          }
        }
      });

      // $("#inherit_parent").click(async (e) => {

      //   if (confirm(`Voulez-vous vraiment supprimer toutes les autorisations uniques et hériter de ses autorisations parent?`)) {
      //     try {

      //       await this.inheritParentPermission(docId, title);

      //     } catch (err) {
      //       alert(err.message);
      //     }
      //   }
      // });

      $("#inherit_parent").click(async (e) => {
        // Disable the button

        if (confirm("Voulez-vous vraiment supprimer toutes les autorisations uniques et hériter de ses autorisations parent?")) {

          $("#inherit_parent").prop("disabled", true);

          // Show loading text or spinner
          $("#inherit_parent").text("Loading...");


          try {

            const list = sp.web.lists.getByTitle("Documents1");
            const mainItem = await list.items.getById(items[0].ID);
            const allSameDocuments = await this.getAllSameDocs(items[0].Title, 'FALSE');

            await mainItem.resetRoleInheritance();
            await Promise.all(allSameDocuments.map(async (item_) => {
              await list.items.getById(item_.ID).resetRoleInheritance();

            }))

              // const { permissions } = await this.getListItemPermissions('https://ncaircalin.sharepoint.com/sites/TestMyGed', "Documents1", itemParent[0].ID, "mgolapkhan.ext@aircalin.nc", "musharaf2897");
              .then(() => {

                alert("Suppression des autorisations uniques et application des autorisations du parents sur ce document.");
                window.location.reload();
              });


            //  await this.inheritParentPermission(docId, title);

            // Reset the button state
            // $("#inherit_parent").prop("disabled", false);
            // $("#inherit_parent").text("Hériter les droits d'accès");


          } catch (err) {


            alert("Erreur lors de l'application de l'autorisation au dossier. Veuillez réessayer.");


            // Enable the button and reset text on error
            $("#inherit_parent").prop("disabled", false);
            $("#inherit_parent").text("Hériter les droits d'accès");
          }
        }
      });

      $("#not_inherit_parent").click(async (e) => {
        // Disable the button

        if (confirm("Voulez-vous vraiment supprimer toutes les autorisations du parents sur ce dossier?")) {

          $("#not_inherit_parent").prop("disabled", true);

          // Show loading text or spinner
          $("#not_inherit_parent").text("Loading...");


          try {

            const list = sp.web.lists.getByTitle("Documents1");
            const mainItem = await list.items.getById(items[0].ID);

            await mainItem.breakRoleInheritance(false, true);
            const allSameDocuments = await this.getAllSameDocs(items[0].Title, 'FALSE');

            await Promise.all(allSameDocuments.map(async (item_) => {
              await list.items.getById(item_.ID).breakRoleInheritance(false, true);
            }))

              .then(() => {

                alert("Attention, l'héritager des drotits pour ce fichier est rompu.");
                window.location.reload();
              });

            //  await this.inheritParentPermission(docId, title);

            // Reset the button state
            // $("#inherit_parent").prop("disabled", false);
            // $("#inherit_parent").text("Hériter les droits d'accès");

          } catch (err) {
            alert("Erreur lors de l'application de l'autorisation au dossier. Veuillez réessayer.");
            // Enable the button and reset text on error
            $("#inherit_parent").prop("disabled", false);
            $("#inherit_parent").text("Hériter les droits d'accès");
          }
        }
      });

      //add_permission_group
      // $("#add_group").click(async (e) => {
      //   if ($("#group_name").val() === "") {
      //     alert("Please select a group.");
      //   }
      //   else {

      //     await this.add_permission_group($("#group_name").val().toString(), items[0].ID, principleIdOfGroup);
      //   }
      // });

      $("#add_group").click(async (e) => {
        // Disable the button


        if ($("#group_name").val() === "") {
          alert("Please select a group.");
        } else {


          const loader = document.createElement("div");
          loader.id = "loader2";
          loader.innerHTML = "<div id='loader-spinner'></div><div id='loader-text'>Veuillez patienter... les autorisations sont en cours d'ajout.</div>"; // Add the text here
          document.body.appendChild(loader);


          $("#add_group").prop("disabled", true);

          // Show loading text or spinner
          $("#add_group").text("Loading...");

          try {

            var selected_permission = $("#permissions_group option:selected").val();

            var permission = 0;

            if (selected_permission === "ALL") {

              permission = 1073741829;
            }

            else if (selected_permission === "READ") {
              permission = 1073741826;

            }
            else if (selected_permission === "READ_WRITE") {
              permission = 1073741830;

            }

            const allSameDocuments = await this.getAllSameDocs(items[0].Title, 'FALSE');

            const list = sp.web.lists.getByTitle("Documents1");

            const _item = await list.items.getById(items[0].ID);
            await _item.breakRoleInheritance(true, false);

            const roleAssignments = _item.roleAssignments;

            await roleAssignments.add(principleIdOfGroup, permission)
              .then(async () => {

                await Promise.all(allSameDocuments.map(async (item_) => {

                  const _item = await list.items.getById(item_.ID);
                  await _item.breakRoleInheritance(true, false);

                  const roleAssignments = _item.roleAssignments;
                  await roleAssignments.add(principleIdOfGroup, permission);

                }));

              });


            document.body.removeChild(loader);
            // await this.add_permission_group(
            //   $("#group_name").val().toString(),
            //   items[0].ID,
            //   principleIdOfGroup
            // );

            alert("Autorisation ajoutée à ce document avec succès.");
            window.location.reload();
            // Reset the button state
            // $("#add_group").prop("disabled", false);
            // $("#add_group").text("Ajouter");
          } catch (error) {

            document.body.removeChild(loader);

            console.error(error);

            // Enable the button and reset text on error
            $("#add_group").prop("disabled", false);
            $("#add_group").text("Ajouter");

            // Optionally, display an error message
            alert("Erreur lors de l'application de l'autorisation au dossier. Veuillez réessayer.");
          }
        }
      });

      //add group notif
      $("#add_group_notif").click(async (e) => {
        this.add_Notif(title, $("#group_name_notif").val(), principleIdOfGroup_notif);
      });

      $("#add_user_notif").click(async (e) => {
        const user: any = await sp.web.siteUsers.getByEmail($("#users_name_notif").val().toString())();
        this.add_Notif(title, user.Title, $("#users_name_notif").val().toString());
      });

    }
  }

  public async removePermissionsOnItem(listTitle, itemId) {
    const list = sp.web.lists.getByTitle(listTitle);
    const item = await list.items.getById(itemId);

    // Break inheritance on the item
    await item.breakRoleInheritance(true, false);

    const roleAssignments = await item.roleAssignments();

    console.log("PERMISSIONS TO B REMOVED", roleAssignments);

    // Delete each role assignment
    for (const roleAssignment of roleAssignments) {
      await item.roleAssignments.getById(roleAssignment.PrincipalId).delete();
      // await item.roleAssignments.remove(roleAssignment.PrincipalId, roleAssignment.);
    }

    console.log('User permissions removed successfully.');
  }

  public async setPermissionsOnItem(listTitle, itemId, permissions) {
    try {
      const list = sp.web.lists.getByTitle(listTitle);
      const item = await list.items.getById(itemId);

      for (const permission of permissions) {

        if (permission.role === "Read") {
          const roleAssignments = item.roleAssignments;
          await roleAssignments.add(permission.id, 1073741826);
        }

        else if (permission.role === "Full Control") {
          const roleAssignments = item.roleAssignments;
          await roleAssignments.add(permission.id, 1073741829);
        }

        else if (permission.role === "Edit") {
          const roleAssignments = item.roleAssignments;
          await roleAssignments.add(permission.id, 1073741830);
        }

        else {

        }
        // Updated line
      }

      console.log('Permissions set successfully.');

      return true;
    } catch (error) {
      console.error('Error setting permissions:', error);
      return false;
    }
  }

  private async getAllSameDocs(docNumber, ifFolder) {
    const all_documents: any[] = await sp.web.lists.getByTitle('Documents1').items
      .select("ID,ParentID,FolderID,Title,revision,IsFolder,description")
      .filter("Title eq '" + docNumber + "' and IsFolder eq '" + ifFolder + "'")
      .get();

    return all_documents;

  }

  private async addBookmark(docID: any, title: any) {
    // Get the current page URL and title
    var url = window.location.href;
    //  var title = document.title;
    let user_current = await sp.web.currentUser();


    // Add the item to the Favourites list
    await sp.web.lists.getByTitle("Marque_Pages").items.add({
      Title: title,
      url: url,
      user: user_current.Title,
      FolderID: docID
    });

    console.log('Item added to Favourites list.');
  }

  private async removeBookmark(docID: any) {
    // Get the current page URL
    var url = window.location.href;

    // Find the item to delete from the Favourites list
    var items = await sp.web.lists.getByTitle("Marque_Pages").items
      .select("ID")
      .filter(`FolderID eq '${Number(docID)}'`)
      .get();

    if (items.length === 0) {
      console.log('Item not found in Favourites list.');
      return;
    }

    // Delete the item from the Favourites list
    await sp.web.lists.getByTitle("Marque_Pages").items.getById(items[0].ID).delete();

    console.log('Item removed from Favourites list.');
  }

  public async checkIfUserIsMemberOfGroup(graphClient: MSGraphClient, groupName: string): Promise<boolean> {
    if (!graphClient) {
      return false;
    }

    try {
      // Get the user's groups
      const groups = await graphClient.api('/me/memberOf')
        .version('v1.0')
        .get();

      // Check if the user is a member of the desired group
      const group = groups.value.find((g: any) => g.displayName === groupName);
      return Boolean(group);
    } catch (error) {
      console.error(error);
      return false;
    }
  }

  public async getUserGroups(graphClient: MSGraphClient): Promise<any[]> {
    if (!graphClient) {
      return [];
    }

    try {
      // Get the user's groups
      const groups = await graphClient.api('/me/memberOf')
        .version('v1.0')
        .get();

      console.log('GROUPS I AM IN : ', groups);
      // Return the array of group objects
      return groups.value;
    } catch (error) {
      console.error(error);
      return [];
    }
  }

  private async updateDocument(folderId: string, title: string, folder: any) {

    let user_current = await sp.web.currentUser();

    const path = $("#path_doc").val().toString();

    if (confirm(`Etes-vous sûr de creer un nouveaux version de ${title} ?`)) {

      const now: Date = new Date();
      const day: number = now.getDate();
      const month: number = now.getMonth() + 1; // Note: Month is zero-indexed
      const year: number = now.getFullYear();

      const date: string = `${day < 10 ? '0' + day : day}/${month < 10 ? '0' + month : month}/${year}`;
      console.log(date); // Output: "10/04/2023"

      try {
        const i = await sp.web.lists.getByTitle('Documents1').items.add({
          Title: $("#input_number").val(),
          description: $("#input_description").val(),
          keywords: $("#input_keywords").val(),
          doc_number: $("#input_number").val(),
          revision: $("#input_revision").val(),
          ParentID: $("#created_by").val(),
          revisionDate: moment($("#input_reviewDate").val()).format("DD/MM/YYYY"),
          IsFolder: "FALSE",
          //createdDate: $("#creation_date").val(),
          // owner: $("#created_by").val(),
          updatedBy: user_current.Title,

          createdDate: $("#creation_date").val().toString(),
          updatedDate: new Date().toLocaleString()

        })
          .then(async (iar) => {

            let listUri: string = '/sites/MyGed/Lists/Documents1';

            // var splitResult = $("#path_doc").toString().split('/');
            // if (splitResult.length >= 6) {
            //   var desiredSubstring = splitResult.slice(5).join('/');
            // }

            await sp.web
              .getFileByServerRelativeUrl(`${listUri}/${iar.data.ID}_.000`)
              .moveTo(`${path}/${iar.data.ID}`);

            const list = sp.web.lists.getByTitle("Documents1");

            let folderId_link = iar.data.ID;
            let title = iar.data.Title;

            await list.items.getById(iar.data.ID).attachmentFiles.add(filename_add, content_add);

            return { list, title, folderId_link };

          })
          .then(async ({ list, title, folderId_link }) => {

            await list.items.getById(folderId_link).update({
              FolderID: parseInt(folderId_link),
              filename: filename_add,
            });

            // await sp.web.lists.getByTitle("InheritParentPermission").items.add({
            //   Title: title,
            //   FolderID: parseInt(folderId_link),
            //   IsDone: "NO",
            //   ParentID: $("#created_by").val(),
            //   IsDoc: "YES"
            // });
            return { title, folderId_link };

          })
          .then(async ({ title, folderId_link }) => {
            await this.createAudit($("#input_number").val(), folderId_link, user_current.Title, "Modification");
            return { title, folderId_link }
          })
          .then(({ title, folderId_link }) => {
            // window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/Document.aspx?document=${title}&documentId=${folderId_link}`;

            alert("Détails mis à jour avec succès");
            return { title, folderId_link };
          })
          .then(({ title, folderId_link }) => {
            //    location.reload(true)
            window.location.href = `${this.context.pageContext.web.absoluteUrl}/SitePages/Document.aspx?document=${title}&documentId=${folderId_link}`;

          });

      }
      catch (err) {
        alert(err.message);
      }


    }
    else {


    }

  }

  private async updateAllDocNumberForDocument(docNumber: any, newDocNumber: any, docid: any) {

    var ifFolder = "FALSE";
    const all_documents: any[] = await sp.web.lists.getByTitle('Documents1').items
      .select("Id,ParentID,FolderID,Title,revision,IsFolder,description")
      .filter("Title eq '" + docNumber + "' and IsFolder eq '" + ifFolder + "'")
      .get();


    try {

      await Promise.all(all_documents.map(async (doc) => {
        const list = sp.web.lists.getByTitle("Documents1");

        // if (item.inheriting !== "NO") {
        const i = await list.items.getById(doc.ID).update({
          Title: newDocNumber.toString(),
        });

      }))
        .then((iar) => {
          alert("Le numéro de document a été modifié pour toutes les versions précédentes.");
          window.location.href = `${this.context.pageContext.web.absoluteUrl}/SitePages/Document.aspx?document=${newDocNumber}&documentId=${docid}`;
        });

    }
    catch (e) {

      alert(e.message);
    }



  }

  private async updateAllDossierForDocument(docNumber: any, docid: any) {

    var ifFolder = "FALSE";
    const all_documents: any[] = await sp.web.lists.getByTitle('Documents1').items
      .select("ID,ParentID,FolderID,Title,revision,IsFolder,description, FileDirRef, UniqueId")
      .filter("Title eq '" + docNumber + "' and IsFolder eq '" + ifFolder + "'")
      .get();


    try {

      await Promise.all(all_documents.map(async (doc) => {
        const list = sp.web.lists.getByTitle("Documents1");
        const path = $("#path_doc").val();

        // if (item.inheriting !== "NO") {
        const i = await list.items.getById(doc.ID).update({
          ParentID: $("#created_by").val(),
        });

        // await sp.web
        // .getFileByServerRelativeUrl(`${doc.FileDirRef}/${doc.ID}_.000`)
        // .moveTo(`${path}/${doc.ID}`);

        await sp.web
          .getFileById(doc.UniqueId)
          .moveTo(`${path}/${doc.ID}`);
        // await sp.web
        // .getFileByServerRelativePath(`${doc.FileDirRef}/${doc.ID}_.000`)
        // .moveTo(`${path}/${doc.ID}`);

      }))
        .then(() => {
          alert("Toutes les versions précédentes ont été déplacées vers le dossier spécifié.");
          window.location.href = `${this.context.pageContext.web.absoluteUrl}/SitePages/Document.aspx?document=${docNumber}&documentId=${docid}`;
        });

    }
    catch (e) {

      alert(e.message);
    }



  }

  private async moveAllRevisionsToArchive(docNumber: any, docid: any) {

    var ifFolder = "FALSE";

    const all_documents: any[] = await sp.web.lists.getByTitle('Documents1').items
      .select("Id,ParentID,FolderID,Title,revision,IsFolder,description,FileDirRef,UniqueId")
      .filter("Title eq '" + docNumber + "' and IsFolder eq '" + ifFolder + "'")
      .get();


    try {

      const list = sp.web.lists.getByTitle("Documents1");

      await Promise.all(all_documents.map(async (doc) => {

        const uri = '/sites/MyGed/Lists/Documents1/1_.000/303_.000';
        // if (item.inheriting !== "NO") {
        const i = await list.items.getById(doc.ID).update({
          ParentID: 791,
        });

        await sp.web
          .getFileById(doc.UniqueId)
          .moveTo(`${uri}/${doc.ID}`);
      }))
        .then(() => {
          alert("Toutes les versions précédentes ont été déplacées vers le dossier archivé.");
          window.location.href = `${this.context.pageContext.web.absoluteUrl}/SitePages/Document.aspx?document=${docNumber}&documentId=${docid}`;
        });

    }
    catch (e) {

      alert(e.message);
    }



  }

  private async createAudit(docTitle: any, folderID: any, userTitle: any, action: any) {


    try {
      // response_same_doc.forEach(async (x) => {

      await sp.web.lists.getByTitle("Audit").items.add({
        Title: docTitle.toString(),
        DateCreated: moment().format("MM/DD/YYYY HH:mm:ss"),
        Action: action.toString(),
        FolderID: folderID.toString(),
        Person: userTitle.toString()
      });
    }

    catch (e) {
      alert("Erreur: " + e.message);
    }

  }

  private async _getAllAudit(id: string, title: string) {

    var value1 = "FALSE";
    var itemID = "";

    var doc_detail: any = await sp.web.lists.getByTitle('Documents1').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + parseInt(id) + "' and IsFolder eq '" + value1 + "'").get();

    await Promise.all(doc_detail.map(async (perm) => {

      itemID = perm.Id;

    }));

    const listContainerDocAudit: Element = this.domElement.querySelector('#splistDocAudit');

    let html: string = `<table id='tbl_doc_audit' class='table table-striped' style="width: 100%;font-size: initial;" >`;

    html += `<thead>
    <tr>
    <th class="text-left">Date</th>
    <th class="text-left">Utilisateur</th>
    <th class="text-left" >Document</th>
    <th class="text-left" >Action</th>
  </tr>
  </thead>
  <tbody id="tbl_documents_audit_bdy">`;

    const allAudit: any[] = await sp.web.lists.getByTitle('Audit').items.select("ID, Title, Person, FolderID, Action, DateCreated").filter("FolderID eq '" + parseInt(id) + "' and Title eq '" + title + "'").getAll();


    console.log("AAAUUDIIIT", allAudit);



    await Promise.all(allAudit.map(async (audit) => {

      html += `
          <tr>
          <td class="text-left">${audit.DateCreated}</td>
          <td class="text-left">${audit.Person}</td>
          <td class="text-left">${audit.Title}</td>
          <td class="text-left">${audit.Action}</td>
         `;

    }))
      .then(() => {

        html += `</tbody>
          </table>`;
        listContainerDocAudit.innerHTML += html;
      });


    const tableSelector = "#tbl_doc_audit"; // Replace with your table selector
    const table = $("#tbl_doc_audit").DataTable();
    //var table = $(tableSelector).DataTable();

    //const table1 = $('#tbl_doc_audit')[0];


    const btnExcelExport = $('#btn_excel');
    const btnPDFExport = $('#btn_pdf');


    btnExcelExport.click(() => {
      // Get all the data in the table, not just the visible data
      const headers = $(table.table().header())
        .find("th")
        .map(function () {
          return [$(this).text()];
        })
        .get();

      const data: any[][] = table.rows().data().toArray();

      console.log("TABLE ROWS LENTGH", data.length);

      data.unshift(headers); // Add the headers at the beginning of the data array

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet1");

      data.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          const worksheetCell = worksheet.getCell(rowIndex + 1, colIndex + 1);
          worksheetCell.value = cell;

          // Apply bold style to the first row
          if (rowIndex === 0) {
            worksheetCell.font = { bold: true };
          }
        });
      });

      workbook.xlsx.writeBuffer().then((buffer) => {
        const blob = new Blob([buffer], { type: "application/octet-stream" });
        saveAs(blob, `${title}_audit.xlsx`);
      });
    });

    btnPDFExport.click(() => {
      // Get the data from the table as an array of arrays

      const data: any[][] = table.rows().data().toArray(); // Create a new jsPDF instance

      const doc = new jsPDF();

      // var dataTableHeaderElements = table.columns().header();

      const headers = $(table.table().header())
        .find("th")
        .map(function () {
          return [$(this).text()];
        })
        .get(); // Add the table data to the PDF document using autoTable

      const date = new Date();
      const formattedDate = date.toLocaleDateString("en-US", {
        day: "numeric",
        month: "long",
        year: "numeric",
      });

      // Add the title
      doc.setFontSize(18);
      doc.text(`${title} audit ` + formattedDate, 10, 10);
      autoTable(doc, {
        head: [headers], // use the first row as the header
        body: data, // use the rest of the rows as the body
        startY: 20, // Adjust the startY value to leave space for the title
      });

      doc.save(`${title}_audit.pdf`);
    });

  }

  private async add_Notif(title, user, email) {

    try {
      // await Promise.all(group_name.map(async (email) => {
      //const user: any = await sp.web.siteUsers.getByEmail(email)();
      await sp.web.lists.getByTitle("toNotify").items.add({
        Title: title,
        groupe_name: user,
        email: email
      });
      // }));

      alert("Notification ajoutée à ce document avec succès.");
      window.location.reload();
    }
    catch (e) {
      alert("Error: " + e.message);
    }


  }

  private async load_folders() {
    const value1 = "TRUE";

    const drp_folders = document.getElementById("folders") as HTMLSelectElement;

    if (!drp_folders) {
      console.error("Dropdown element not found");
      return;
    }

    const all_folders = await sp.web.lists.getByTitle('Documents1').items
      .select("ID,ParentID,FolderID,Title,IsFolder,description,FileDirRef")
      .top(5000)
      .filter("IsFolder eq '" + value1 + "'")
      .get();

    console.log("ALL FOLDERS", all_folders.length);


    await Promise.all(all_folders.map(async (result) => {

      const opt = document.createElement('option');

      opt.value = result.Title + '_' + result.FolderID; // Concatenate FolderID and Title to create a unique value      // opt.text = result.Title;
      opt.setAttribute('data-value', result.FolderID);
      opt.setAttribute('data-filedirref', `${result.FileDirRef}/${result.ID}_.000`)
      opt.dataset// Set the value of the option to the folder ID

      // create a span element for the description
      const description = document.createElement('span');
      description.innerHTML = result.description;

      // add a line break element before the description
      const lineBreak = document.createElement('br');
      opt.appendChild(lineBreak);

      // opt.insertAdjacentElement('beforeend', description);

      opt.appendChild(description);
      drp_folders.appendChild(opt);

    }));


  }

  private fileUpload() {

  }

  private createPath(listDoc: any) {

    const listContainerDocPath: Element = this.domElement.querySelector('#doc_path');
    let html: string = `<ul class="breadcrumb" id="breadcrumb">`;


    listDoc.forEach((item) => {

      if (item.parentTitle_doc !== undefined) {
        html += `<li><a href="${this.context.pageContext.web.absoluteUrl}/SitePages/documentation.aspx?folder=${item.parentId_doc}">${item.parentTitle_doc}</a></li>`;
      }
    });

    html += `</ul>`;

    listContainerDocPath.innerHTML += html;

  }

  private getFileExtensionFromUrl(url: string): string {
    const lastDotIndex = url.lastIndexOf('.');
    if (lastDotIndex === -1) {
      // no dot in the URL
      return '';
    }

    const pathAfterLastSlash = url.slice(url.lastIndexOf('/') + 1);
    const lastSlashIndex = pathAfterLastSlash.lastIndexOf('/');
    const filenameWithExtension = pathAfterLastSlash.slice(lastSlashIndex + 1);
    const extension = filenameWithExtension.slice(filenameWithExtension.lastIndexOf('.') + 1);

    return extension;
  }

  private async _getDocDetails(id: number, title: any) {

    const parentArray = await this.getParentID(this.getDocId(), this.getDocTitle());

    if (parentArray.parentTitle !== null) {
      this.createPath(parentArray.parentTitle);
    }

    console.log("PARENT BREADCRUMB", parentArray);



    var urlFile_download = '';
    var pdfNameDownload = '';


    $('#file_ammendment_update').on('change', () => {
      const input = document.getElementById('file_ammendment_update') as HTMLInputElement | null;


      var file = input.files[0];
      var reader = new FileReader();

      reader.onload = ((file1) => {
        return (e) => {
          console.log(file1.name);

          filename_add = file1.name,
            content_add = e.target.result
          $("#input_filename").val(file1.name);
        };
      })(file);

      reader.readAsArrayBuffer(file);
    });

    // Get the document details
    const itemDoc: any[] = await sp.web.lists.getByTitle('Documents1').items
      .select("Id,ParentID,FolderID,Title,revision,IsFolder, description, revisionDate, keywords, owner, updatedBy, updatedDate, createdDate, attachmentUrl, IsFiligrane, IsDownloadable, filename, FileDirRef, UniqueId, source")
      .filter(`FolderID eq '${id}' and Title eq '${title}' and IsFolder eq 'FALSE'`)
      .get();

    // if (itemDoc[0].source === "qpulse" || itemDoc[0].source === "gedeon1") {
    //   $("#input_reviewDate").prop("disabled", true);
    // }

    if (itemDoc[0].ParentID == 0) {

      $("#input_type_doc").val("");
      $("#path_doc").val("");

    }


    else {
      var itemFolder: any[] = await sp.web.lists.getByTitle('Documents1').items
        .select("Id,ParentID,FolderID,Title, FileDirRef")
        .filter("FolderID eq '" + parseInt(itemDoc[0].ParentID) + "' and IsFolder eq 'TRUE'")
        .get();

      $("#input_type_doc").val(itemFolder[0].Title);
      $("#path_doc").val(itemDoc[0].FileDirRef);

    }


    // Get the folder details
    await sp.web.lists.getByTitle("Documents1")
      .items
      .getById(parseInt(itemDoc[0].Id))
      .attachmentFiles
      .select('FileName', 'ServerRelativeUrl')
      .get()
      .then(responseAttachments => {
        responseAttachments
          .forEach(attachmentItem => {

            pdfNameDownload = attachmentItem.FileName;
            urlFile_download = attachmentItem.ServerRelativeUrl;
          });
      });

    var checkbox_fili = document.getElementById("check1") as HTMLInputElement;
    var checkbox_download = document.getElementById("check2") as HTMLInputElement;

    if (itemDoc[0].IsFiligrane == "YES") {
      checkbox_fili.checked = true;
    }
    else {
      checkbox_fili.checked = false;
    }

    if (itemDoc[0].IsDownloadable == "YES") {
      checkbox_download.checked = true;
    }
    else {
      checkbox_download.checked = false;
    }


    var condition1 = /^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}$/.test(itemDoc[0].createdDate);

    var condition2 = /^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}$/.test(itemDoc[0].createdDate);

    var condition3 = /^(\d{2}\/\d{2}\/\d{4}), (\d{2}:\d{2}:\d{2})$/.test(itemDoc[0].createdDate);

    var condition4 = /^(\d{2})\/(\d{2})\/(\d{4}), (\d{2}):(\d{2}):(\d{2}) (AM|PM)$/.test(itemDoc[0].createdDate);

    var condition5 = /^(\d{2})\/(\d{2})\/(\d{4})$/.test(itemDoc[0].createdDate);



    var condition11 = /^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}$/.test(itemDoc[0].revisionDate);

    var condition22 = /^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}$/.test(itemDoc[0].revisionDate);

    var condition33 = /^(\d{2}\/\d{2}\/\d{4}), (\d{2}:\d{2}:\d{2})$/.test(itemDoc[0].revisionDate);

    var condition44 = /^(\d{2})\/(\d{2})\/(\d{4}), (\d{2}):(\d{2}):(\d{2}) (AM|PM)$/.test(itemDoc[0].revisionDate);

    var condition55 = /^(\d{2})\/(\d{2})\/(\d{4})$/.test(itemDoc[0].revisionDate);


    var dateRevision = itemDoc[0].revisionDate;


if (dateRevision !== null && (condition11 || condition22)) {
      // var dateString = itemDoc[0].createdDate;

      var dateParts2 = dateRevision.split(' ');

      var dateParts = dateParts2[0].split("/");
      var year = dateParts[2];
      var month = dateParts[1];
      var day = dateParts[0];
      dateRevision = year + "-" + month + "-" + day;

    }

    else if (dateRevision !== null && condition33) {

      var dateParts2 = dateRevision.split(',');

      var dateParts = dateParts2[0].split("/");
      var year = dateParts[2];
      var month = dateParts[1];
      var day = dateParts[0];

      dateRevision = year + "-" + month + "-" + day;

    }

    else if (dateRevision !== null && condition44) {

      var dateParts2 = dateRevision.split(',');

      var dateParts = dateParts2[0].split("/");
      var year = dateParts[2];
      var month = dateParts[1];
      var day = dateParts[0];
      dateRevision = year + "-" + month + "-" + day;

    }

    else if (dateRevision !== null && condition55) {

      var dateParts2 = dateRevision.split(',');

      var dateParts = dateParts2[0].split("/");
      var year = dateParts[2];
      var month = dateParts[1];
      var day = dateParts[0];
      dateRevision = year + "-" + month + "-" + day;

    }


    var dateCreated = itemDoc[0].createdDate;

    if (condition1 || condition2) {
      // var dateString = itemDoc[0].createdDate;

      var dateParts2 = dateCreated.split(' ');

      var dateParts = dateParts2[0].split("/");
      var year = dateParts[2];
      var month = dateParts[1];
      var day = dateParts[0];
      dateCreated = year + "-" + month + "-" + day;

    }

    else if (condition3) {

      var dateParts2 = dateCreated.split(',');
      var dateParts = dateParts2[0].split("/");
      var year = dateParts[2];
      var month = dateParts[1];
      var day = dateParts[0];
      dateCreated = year + "-" + month + "-" + day;

    }

    else if (condition4) {

      var dateParts2 = dateCreated.split(',');
      var dateParts = dateParts2[0].split("/");
      var year = dateParts[2];
      var month = dateParts[1];
      var day = dateParts[0];

      dateCreated = year + "-" + month + "-" + day;

    }
    else if (condition5) {

      var dateParts2 = dateCreated.split(',');
      var dateParts = dateParts2[0].split("/");
      var year = dateParts[2];
      var month = dateParts[1];
      var day = dateParts[0];

      dateCreated = year + "-" + month + "-" + day;

    }


    // Set the document details in the UI
    // $("#input_type_doc").val(itemFolder[0].Title);
    $("#input_number").val(itemDoc[0].Title);
    $("#input_revision").val(itemDoc[0].revision);
    $("#input_keywords").val(itemDoc[0].keywords);
    $("#input_description").val(itemDoc[0].description);
    $("#created_by").val(itemDoc[0].ParentID);
    $("#updated_by").val(itemDoc[0].updatedBy);
    $("#updated_time").val(itemDoc[0].updatedDate);
    $("#creation_date").val(dateCreated);
    $("#h2_doc_title").text(itemDoc[0].Title);
    // $("#path_doc").val(itemDoc[0].FileDirRef);

    $("#input_reviewDate").val(dateRevision);

    // if (itemDoc[0].revisionDate) {
    //   $("#input_reviewDate").val((itemDoc[0].revisionDate).split(" ")[0]);
    // }

    // else {
    //   var dateString = itemDoc[0].revisionDate.toString();
    //   var dateParts = dateString.split(" ")[0].split("/");
    //   var convertedDate = dateParts[0] + "/" + dateParts[1] + "/" + dateParts[2];
    //   $("#input_reviewDate").val(convertedDate);
    // }


    // Set the URL for downloading the attachment
    let url = '';
    let externalUrl = itemDoc[0].attachmentUrl;

    if (externalUrl == undefined || externalUrl == null || externalUrl == "") {
      url = urlFile_download;
    } else {
      url = externalUrl;
    }

    if (itemDoc[0].IsDownloadable == "NO") {
      $("#download_doc").css("display", "none");
    }

    // var xx = await this.getPermissionLevel(this.graphClient, "TestMyGed", "Documents1", itemDoc[0].Id);
    // console.log("GRAPH PERMISSION", xx);

    // Open the attachment in a new tab

    let user_current = await sp.web.currentUser();
    $("#open_doc").click(async (e) => {

      //   url = "https://ncaircalin.sharepoint.com" + url;
      if (this.getFileExtensionFromUrl(url) !== "pdf") {
        //  if (itemDoc[0].IsFiligrane === "NO") {
        // window.open(`${url}`, '_blank');
        await this.downloadFile(url, itemDoc[0].Title + "." + this.getFileExtensionFromUrl(url) , itemDoc[0].FolderID);
      }

      else {

        // create a div element to blur the screen
        const blurDiv = document.createElement('div');
        blurDiv.classList.add('blur');
        document.body.appendChild(blurDiv);

        // // create a div element to show the loader
        const loaderDiv = document.createElement('div');
        loaderDiv.classList.add('loader1');
        document.body.appendChild(loaderDiv);

        try {
          //await this.openPDFInIframe(url, 'UNCONTROLLED COPY - Downloaded on: ');
          // window.open(`${url}`, '_blank');

          await this.createWebpageInNewTab(url, itemDoc[0].filename);
          // await this.createWebpageInIframe2(url);
          // await this.createWebpageInNewTab2(url);

        } finally {
          // remove the loader and the blur elements
          document.body.removeChild(loaderDiv);
          document.body.removeChild(blurDiv);
        }
      }
      // }

      // else {
      //   alert("This is not a pdf file");

      //   window.open(`${url}`, '_blank');
      // }

      await this.createAudit(itemDoc[0].Title, itemDoc[0].FolderID, user_current.Title, "Consultation");
    });


    //check user permission
    // var userPermission = this.getUserPermissionLevelOnSharePointListItem("Documents1", itemDoc[0].Id, user_current.Email);
    // console.log("User Permission", userPermission);


    // Delete the document
    $("#delete_doc").click(async (e) => {
      //  if (confirm(`Are you sure you want to delete ${itemDoc[0].Title}?`)) {
      if (confirm(`Voulez-vous vraiment archiver ce document et toutes les révisions précédentes : ${itemDoc[0].Title} ?`)) {
        try {

          await this.moveAllRevisionsToArchive(itemDoc[0].Title, itemDoc[0].FolderID);
          // await sp.web.lists.getByTitle('Documents1').items.getById(parseInt(itemDoc[0].Id)).recycle();
          // alert("Document deleted successfully.");
          // window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${itemDoc[0].ParentID}`;
        } catch (err) {
          alert(err.message);
        }
      }
    });

    //archive doc
    $("#archive_btn").click(async (e) => {

      if (confirm(`Voulez-vous vraiment archiver ce document : ${itemDoc[0].Title} ?`)) {
        try {

          const list = sp.web.lists.getByTitle("Documents1");

          const uri = '/sites/MyGed/Lists/Documents1/1_.000/303_.000';

          const i = await list.items.getById(itemDoc[0].ID).update({
            ParentID: 791,
          })
            .then(async () => {

              await sp.web
                .getFileById(itemDoc[0].UniqueId)
                .moveTo(`${uri}/${itemDoc[0].ID}`);

              alert("Le document a été archivé.");
              window.location.href = `${this.context.pageContext.web.absoluteUrl}/SitePages/Document.aspx?document=${itemDoc[0].Title}&documentId=${itemDoc[0].FolderID}`;
            });
        } catch (err) {
          alert(err.message);
        }
      }

    });

    $("#download_doc").click(async (e) => {

      const specialCharsRegex = /[^\w\s]/g; // Regex to match special characters excluding dot and forward slash

      const cleanedTitle = itemDoc[0].Title.replace(specialCharsRegex, "_"); // Replace special characters with an underscore (_)
      const finalTitle = cleanedTitle.replace(/\./g, "_").replace(/\//g, "_");

      if (this.getFileExtensionFromUrl(url) === "pdf") {

        if (itemDoc[0].IsFiligrane === "NO") {
          await this.downloadDocWithoutFili(url, finalTitle, itemDoc[0].FolderID);
        } else {
          await this.downloadDoc(url, finalTitle, itemDoc[0].FolderID, 'UNCONTROLLED COPY - Downloaded on ');
        }

      }

      else {
        await this.downloadFile(url, finalTitle + "." + this.getFileExtensionFromUrl(url) , itemDoc[0].FolderID);
      }


    });


    $("#update_details_doc").click(async (e) => {
      const fileInput = document.getElementById("file_ammendment_update") as HTMLInputElement;

      // if (($("#input_type_doc").val() === itemFolder[0].Title ||
      //   $("#input_number").val() === itemDoc[0].Title ||
      //   $("#input_keywords").val() === itemDoc[0].keywords ||
      //   $("#input_description").val() === itemDoc[0].description) && $("#input_revision").val() !== itemDoc[0].revision && fileInput.value) {
      // The fields are unchanged, so update the document metadata
      //  const folder = await this.getFolder(itemFolder[0].FolderID);
      // const fileInput = document.getElementById("myFileInput") as HTMLInputElement;

      //   if ($("#myFileInput").get(0).files.length === 0) 
      if ($("#input_revision").val() !== itemDoc[0].revision && fileInput.value) {

        if (fileInput.files.length == 0 || $("#input_revision").val() === "") {
          alert("Veuillez entrer le numéro de révision et télécharger votre document.");
        }

        else {

          var confirmed = confirm("Êtes-vous sûr de vouloir créer une nouvelle version de ce document ?");
          if (confirmed) {

            // User clicked OK
            // Your delete code goes here
            await this.updateDocument(itemDoc[0].FolderID, itemDoc[0].Title, Number($("#input_type_doc").val()));
            alert("Vous avez créé une nouvelle version de : " + itemDoc[0].Title);

          } else {
            // User clicked Cancel
            // Your cancel code goes here
          }

        }

        // await this.updateDocument(itemDoc[0].FolderID, itemDoc[0].Title, folder.ID);

      }

      else if ($("#input_number").val() !== itemDoc[0].Title) {

        var confirmed = confirm("Êtes-vous sûr de vouloir appliquer cette modification à tous les Documents1 portant ce numéro de document?");
        if (confirmed) {

          await this.updateAllDocNumberForDocument(itemDoc[0].Title, $("#input_number").val(), itemDoc[0].FolderID);
        }
        else {

        }

      }

      else if (Number($("#created_by").val()) !== itemDoc[0].ParentID) {

        // alert($("#created_by").val());
        // alert(itemDoc[0].ParentID);

        var confirmed = confirm("Êtes-vous sûr de bouger les Documents1 portant ce numéro de document?");
        if (confirmed) {

          await this.updateAllDossierForDocument(itemDoc[0].Title, itemDoc[0].FolderID);
        }
        else {

        }

      }

      else {

        var confirmed = confirm("Êtes-vous sûr de vouloir mettre à jour les détails de ce document?");
        if (confirmed) {
          // User clicked OK
          // Your delete code goes here
          await this.updateDocMetadata(itemDoc[0].Id, itemDoc[0].FolderID, user_current.Title);

          alert("Vous avez modifié certaines métadonnées de : " + itemDoc[0].Title);

        } else {
          // User clicked Cancel
          // Your cancel code goes here
        }
        // The fields are changed, so get the folder details and update the document metadata
        // const folder = await this.getFolder(itemFolder[0].FolderID);
        // await this.updateDocMetadata(itemDoc[0].Id, folder.ID, itemDoc[0].FolderID, user_current.Title);

      }
      // await this.createAudit()

    });

    $("#edit_cancel_doc").click(async (e) => {

      window.location.href = `${this.context.pageContext.web.absoluteUrl}/SitePages/documentation.aspx?folder=${itemFolder[0].FolderID}`;

    });


    const bookmarkSwitch = document.getElementById("bookmark-switch") as HTMLInputElement;
    bookmarkSwitch.addEventListener("change", async () => {
      await this.handleBookmarkSwitchChange.call(this, bookmarkSwitch.checked, itemDoc[0].FolderID, itemDoc[0].Title);
    });

    const user: ISiteUserInfo = await sp.web.currentUser();

    // await this.checkUserPermissions( "Documents1", itemDoc[0].ID , user.Id);

    var items = await sp.web.lists.getByTitle("Marque_Pages").items
      .select("ID")
      .filter(`FolderID eq '${itemDoc[0].FolderID}' and user eq '${user.Title}'`)
      .get();

    if (items.length === 0) {
      console.log('Item not found in Favourites list.');
      return this.setBookmarkSwitchState(false);
    }

    else {

      return this.setBookmarkSwitchState(true);

    }

  }

  public async downloadFile(url, filename, folderID) {

    try {

      const message = document.createElement('div');
      message.textContent = 'Downloading file...';
      message.style.cssText = `
        position: fixed;
        bottom: 0;
        width: 100%;
        background-color: #f54630;
        color: white;
        text-align: center;
        font-size: 24px;
        padding: 10px 0;
      `;

      document.body.appendChild(message);

      const response = await fetch(url);
      const blob = await response.blob();
      const a = document.createElement("a");
      const downloadUrl = URL.createObjectURL(blob);
      a.href = downloadUrl;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(downloadUrl);


      document.body.removeChild(message);
    }

    catch (e) {
      alert("Cannot download this file for the following reason: " + e);

      // Remove the message from the DOM in case of an error
      const message = document.querySelector('div');
      if (message) {
        document.body.removeChild(message);
      }

      window.location.reload();
    }

  }

  private async createWebpageInNewTab(url, filename: any) {
    try {

      const pdfBytes = await this.generatePdfBytes2(url);
      const pdfUrl = URL.createObjectURL(new Blob([pdfBytes], { type: 'application/pdf' }));


      const overlay = document.createElement('div');
      overlay.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-color: rgba(0, 0, 0, 0.5);
        display: flex;
        justify-content: center;
        align-items: center;
      `;

      const closeButton = document.createElement('button');
      closeButton.textContent = 'Close';
      closeButton.style.cssText = `
      position: absolute;
      top: 80px;
      left: 27%;
      padding: 5px 10px;
      background-color: #fa0f00;
      border: none;
      cursor: pointer;
      font-size: large;
      color: white;
      `;

      //     const mediaQuery = window.matchMedia('(max-width: 600px)');
      //     if (mediaQuery.matches) {
      //       closeButton.style.cssText += `
      //       position: absolute;
      //       bottom: 3px;
      //       right: 19px;
      //       padding: 5px 10px;
      //       background-color: darkblue;
      //       border: none;
      //       cursor: pointer;
      //       font-size: smaller;
      //       color: white;
      // `;
      //     }

      overlay.appendChild(closeButton);

      const iframe = document.createElement('iframe');
      iframe.style.cssText = `
        width: 39%;
        height: 97%;
      `;

      overlay.appendChild(iframe);
      document.body.appendChild(overlay);

      closeButton.addEventListener('click', function () {
        document.body.removeChild(overlay);
      });

      const iframeDoc = iframe.contentDocument || iframe.contentWindow.document;

      iframeDoc.open();
      iframeDoc.write(`
        <html>
          <head>
            <title>MyGed PDF Viewer</title>
            <style>
              body { margin: 0; }
              #adobe-dc-view { width: 100%; height: 100%; }
            </style>
            <script src="https://acrobatservices.adobe.com/view-sdk/viewer.js"></script>
            <script type="text/javascript">
            document.addEventListener("adobe_dc_view_sdk.ready", function(){ 
              var adobeDCView = new AdobeDC.View({clientId: "d77e60caf95e49169e0443eb71689bd5", divId: "adobe-dc-view"});
          
              function convertPDFToBlob(url) {
                return new Promise((resolve, reject) => {
                  fetch(url)
                    .then(response => {
                      if (!response.ok) {
                        throw new Error('Failed to fetch the PDF file.');
                      }
                      return response.blob();
                    })
                    .then(blob => {
                      resolve(blob);
                    })
                    .catch(error => {
                      reject(error);
                    });
                });
              }
          
              // var url = 'https://example.com/path/to/file.pdf';
              // var filename = 'example.pdf';
          
              // var x = convertPDFToBlob('${url}');
              // x.then(blob => {
                adobeDCView.previewFile({
                  content:{location: {url: '${pdfUrl}'}},
                  metaData: { fileName: "${filename}" }
                }, { embedMode: "IN_LINE", showDownloadPDF: false, showPrintPDF: false});
              });
            // });
          </script>
          </head>
          <body>
            <div id="adobe-dc-view"></div>
          </body>
        </html>
      `);
      iframeDoc.close();
    } catch (error) {
      console.error('Error:', error);
    }
  }

  public async getPermissionLevel(graphClient: MSGraphClient, siteName: string, listName: string, itemId: string): Promise<string> {
    try {
      // Get site ID
      const siteResponse = await graphClient.api(`/sites/${siteName}`).get();
      const siteId = siteResponse.id;

      // Get list ID
      const listResponse = await graphClient.api(`/sites/${siteId}/lists?$filter=displayName eq '${listName}'`).get();
      const listId = listResponse.value[0].id;

      // Get item permissions
      const response = await graphClient.api(`/sites/${siteId}/lists/${listId}/items/${itemId}/driveItem/permissions`).get();
      const permissions = response.value;

      // Get current user
      const currentUser = await graphClient.api('/me').get();
      const currentUserObjectId = currentUser.id;

      // Get permission level for current user
      const permissionLevel = permissions.find(p => p.grantedToIdentities.some(identity => identity.user && identity.user.id === currentUserObjectId));

      if (permissionLevel) {
        switch (permissionLevel.role) {
          case 'write':
            return 'Contribute';
          case 'read':
            return 'Read';
          case 'edit':
            return 'Edit';
          case 'owner':
            return 'Full Control';
          default:
            return 'Unknown';
        }
      } else {
        throw new Error(`User does not have permissions on item ${itemId}`);
      }
    } catch (error) {
      console.log(error);
      throw error;
    }
  }

  private async setBookmarkSwitchState(isChecked: boolean) {
    const bookmarkSwitch = document.getElementById("bookmark-switch") as HTMLInputElement;


    bookmarkSwitch.checked = isChecked;

    if (isChecked) {
      const starIcon = document.querySelector(".star-icon") as HTMLElement;

      starIcon.classList.remove("fa-regular");
      starIcon.classList.add("fa-solid");
    }
  }

  private async handleBookmarkSwitchChange(isChecked: boolean, doc_id: any, title: any) {
    try {
      if (isChecked) {
        await this.addBookmark(Number(doc_id), title);
        alert("Vous avez ajouté ce document comme favori.");
        window.location.reload();
      } else {
        await this.removeBookmark(doc_id);
        alert("Vous avez supprimé ce document des favoris.");
        window.location.reload();
      }
    } catch (error) {
      console.error("Failed to update bookmark:", error);
    }
  }

  private async updateDocMetadata(id: any, folderID: any, userTitle: any,) {

    const path = $("#path_doc").val().toString();

    const item: any = await sp.web.lists.getByTitle("Documents1").items
      .getById(id)
      .select("ID,ParentID,FolderID,Title,revision,IsFolder,description,FileDirRef,UniqueId")
      .get();


    var checkbox_fili = document.getElementById("check1") as HTMLInputElement;
    var filivalue = "NO";

    var checkbox_download = document.getElementById("check2") as HTMLInputElement;
    var download_value = "NO";

    if (checkbox_fili.checked === true) {
      filivalue = "YES";
    } else {
      filivalue = "NO";
    }

    if (checkbox_download.checked === true) {
      download_value = "YES";
    } else {
      download_value = "NO";
    }


    try {
      const list = sp.web.lists.getByTitle("Documents1");

      const i = await list.items.getById(id).update({
        // Title: $("#input_number").val(),
        description: $("#input_description").val(),
        keywords: $("#input_keywords").val(),
        doc_number: $("#input_number").val(),
        ParentID: $("#created_by").val(),
        IsFiligrane: filivalue,
        revisionDate: moment($("#input_reviewDate").val()).format("DD/MM/YYYY"),
        IsDownloadable: download_value,
        createdDate: moment($("#creation_date").val()).format("DD/MM/YYYY")
      })
        .then(async () => {

          let listUri: string = '/sites/MyGed/Lists/Documents1';

          if (path === "") {
            await sp.web
              .getFileByServerRelativeUrl(`${listUri}/${id}_.000`)
              .moveTo(`${path}/${id}`);

            // const i = await list.items.getById(id).update({
            //   Current: path
            // });

          }

          else if (path !== item.FileDirRef) {

            await sp.web
              .getFileByServerRelativeUrl(`${item.FileDirRef}/${id}_.000`)
              .moveTo(`${path}/${id}`);

            const i = await list.items.getById(id).update({
              Current: path
            });
          }

          else {

          }

        })
        .then(async () => {
          await this.createAudit($("#input_number").val(), folderID, userTitle, "Modification");
        })
        .then(() => {

          alert("Détails mis à jour avec succès");
        })
        .then(() => {
          window.location.reload();
        });
    }

    catch (err) {
      alert(err.message);
    }

  }

  private async _getAllVersions(title: string, doc_id: any) {

    {
      //display so table
      // $("#table_version_doc").css("display", "block");
      const listContainerDocVersions: Element = this.domElement.querySelector('#splistDocVersions');
      let html: string = `<table id='tbl_doc_versions' class='table table-striped' style="width: 100%;font-size: initial;" >`;

      html += `<thead>
      <tr>
        <th class="text-left">ID</th>
        <th class="text-left">Nom du document</th>
        <th class="text-left" >Url</th>
        <th class="text-left" >Revision</th>
        <th class="text-left" >Description</th>
        <th class="text-left" >Pdf Name</th>
        <th class="text-left" >Folder ID</th>
        <th class="text-left" >IsFiligrane</th>
        <th class="text-left" >Imprimable</th>
        <th class="text-center" >Actions</th>
      </tr>
    </thead>
    <tbody id="tbl_documents_versions_bdy">`;


      var value1 = "FALSE";


      const all_documents_versions: any[] = await sp.web.lists.getByTitle('Documents1').items
        .select("Id,ParentID,FolderID,Title,revision,IsFolder,description, attachmentUrl, IsFiligrane, IsDownloadable, filename")
        .filter("Title eq '" + title + "' and IsFolder eq '" + value1 + "' and FolderID ne '" + doc_id + "'")
        .get();

      // First, sort the array by revision in descending order
      // all_documents_versions.sort((a, b) => b.revision - a.revision);

      // Then, find the highest revision
      //const highestRevision = all_documents_versions[0].revision;

      var filtered_docs_versions = all_documents_versions;

      // Finally, filter the array and remove the items with the highest revision
      //  const filtered_documents_versions_1 = all_documents_versions.filter((document) => document.revision < highestRevision);

      //  const filtered_documents_versions_2 = all_documents_versions.filter((document) => document[0].revision !== null && document[0].FolderID !== doc_id);

      const filtered_documents_versions_2 = filtered_docs_versions.filter((document) => {
        return document.revision !== null && document.FolderID !== doc_id;
      });

      //const filteredItems = filtered_documents_versions_2.filter(doc => doc.FolderID !== doc_id);


      if (filtered_documents_versions_2.length > 0) {
        $("#table_version_doc").css("display", "block");

        //  await Promise.all(contract.map(async (result) => {
        // await response_doc_versions.forEach(async (element_version) => {

        await Promise.all(filtered_documents_versions_2.map(async (element_version) => {

          var pdfName = '';
          var urlFile = '';
          var url = '';
          var attachmentUrl = element_version.attachmentUrl;

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

          var x = element_version.IsDownloadable;
          var y = element_version.Id;


          html += `
          <tr>
          <td class="text-left">${element_version.Id}</td>

          <td class="text-left">${element_version.Title}</td>

          <td class="text-left"> 
         ${url}          
          </td>

          <td class="text-left"> 
          ${element_version.revision}          
          </td>

          <td class="text-left"> 
          ${element_version.description}          
          </td>

          <td class="text-left"> 
          ${pdfName}          
          </td>

          <td class="text-left"> 
          ${element_version.FolderID}          
          </td>

          <td class="text-left"> 
          ${element_version.IsFiligrane}          
          </td>

          <td class="text-left"> 
          ${element_version.IsDownloadable}          
          </td>
          
          <td class="text-center">

         <a href="#"  title="Voir le document" id="${element_version.Id}_view_doc_version" class="btn_view_doc" style="padding-left: inherit;">
         <i class="fa-sharp fa-solid fa-eye" style="font-size: x-large;"></i>
         </a>

         <a href="#" title="Telecharger le document" title="imprimer le document" id="${element_version.Id}_download_doc" class="btn_download_doc" style="padding-left: 1em;">
         <i class="fa-solid fa-download" style="font-size: x-large;"></i>

     
         </a>

          </td>

         `;

          return { x, y };

        }))
          .then((results: { x: any; y: any; }[]) => {
            if (results[0].x === "NO") {
              $(`#${results[0].y}_download_doc`).css("display", "none");
            }

          })
          .then(() => {

            html += `</tbody>
          </table>`;
            listContainerDocVersions.innerHTML += html;
          });



        table = $('#tbl_doc_versions').DataTable({
          columnDefs: [{
            target: 0,
            visible: false,
            searchable: false
          },
          {
            target: 2,
            visible: false,
            searchable: false
          }
            ,
          {
            target: 5,
            visible: false,
            searchable: false
          }
            ,
          {
            target: 6,
            visible: false,
            searchable: false
          }
            ,
          {
            target: 7,
            visible: false,
            searchable: false
          }
            ,
          {
            target: 8,
            visible: false,
            searchable: false
          }
          ]
        });

        table.rows().every(function () {
          // Get the data in the eighth column of the current row
          var data = this.cell(this.index(), 8).data();

          // Check if the value is "NO"
          if (data === "NO") {
            // Get the button element in the current row
            var btnCurrentRow = $(this.node()).find(".btn_download_doc");
            // Hide the button element in the current row
            btnCurrentRow.css('display', 'none');
          }
        });


        $('#tbl_doc_versions tbody').on('click', '.btn_view_doc', async (event) => {
          var data = table.row($(event.currentTarget).parents('tr')).data();



          if (this.getFileExtensionFromUrl(data[2]) !== "pdf") {
            // alert("FILIGRANE = NO");
            window.open(`${data[2]}`, '_blank');
          }

          else {


            // create a div element to blur the screen
            const blurDiv = document.createElement('div');
            blurDiv.classList.add('blur');
            document.body.appendChild(blurDiv);

            // // create a div element to show the loader
            const loaderDiv = document.createElement('div');
            loaderDiv.classList.add('loader1');
            document.body.appendChild(loaderDiv);

            try {
              // await this.openPDFInIframe(data[2], 'ARCHIVED COPY - Downloaded on: ');
              //await this.createWebpageInIframe2(data[2]);

              // window.open(data[2], '_blank');
              await this.createWebpageInNewTab(data[2], data[5]);
            } finally {
              // remove the loader and the blur elements
              document.body.removeChild(loaderDiv);
              document.body.removeChild(blurDiv);

            }
          }
          //  window.open(`${data[2]}`, '_blank');
        });


        $('#tbl_doc_versions tbody').on('click', '.btn_download_doc', async (event) => {
          var data = table.row($(event.currentTarget).parents('tr')).data();

          const specialCharsRegex = /[^\w\s]/g; // Regex to match special characters excluding dot and forward slash
          const cleanedTitle = data[1].replace(specialCharsRegex, "_");
          const finalTitle = cleanedTitle.replace(/\./g, "_").replace(/\//g, "_");
          const finalRevision = data[3].replace(/\./g, "_").replace(/\//g, "_");


          if (data[7] === "NO") {
            await this.downloadDocWithoutFili(data[2], `${finalTitle}_${finalRevision}`, data[6]);
          }
          else {
            await this.downloadDoc(data[2], `${finalTitle}_${finalRevision}`, data[6], 'ARCHIVED COPY - Downloaded on ');
          }
        });

      }

    }

  }

  public async checkIfUserIsAdmin(graphClient: MSGraphClient): Promise<boolean> {
    try {
      const groups = await graphClient.api('/me/transitiveMemberOf/microsoft.graph.group?$count=true&$top=999').get();
      const groupList = groups.value;
      const isAdmin = groupList.some(group => group.displayName === 'MYGED_ADMIN');
      const isRefUser = groupList.some(group => group.displayName.startsWith('MYGED_REF'));
      const isGuestUser = groupList.some(group => group.displayName.startsWith('MYGED_GUEST'));
      return isAdmin || isRefUser || isGuestUser;
    } catch (error) {
      console.log(error);
      return false;
    }
  }

  private async _getAllNotifications(id: string) {

    var value1 = "FALSE";
    var itemID = "";
    var folderID = "";
    var docTitle = this.getDocTitle();

    var doc_detail: any = await sp.web.lists.getByTitle('Documents1').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + parseInt(id) + "' and IsFolder eq '" + value1 + "'").get();

    await Promise.all(doc_detail.map(async (perm) => {

      itemID = perm.Id;
      folderID = perm.FolderID;

    }));

    const listContainerDocNotif: Element = this.domElement.querySelector('#splistDocNotifications');

    let html: string = `<table id='tbl_doc_notif' class='table table-striped' style="width: 100%;font-size: initial;" >`;

    html += `<thead>
    <tr>
      <th class="text-left">ID</th>
      <th class="text-left">Nom</th>

      <th class="text-left" >Actions</th>
    </tr>
  </thead>
  <tbody id="tbl_documents_notif_bdy">`;

    // const allNotif: any[] = await sp.web.lists.getByTitle('Notifications').items.select("ID, Title, group_person, revisionDate, toNotify, webLink, description, FolderID, LoginName").filter("FolderID eq '" + folderID.toString() + "'").getAll();
    const allNotif: any[] = await sp.web.lists.getByTitle('toNotify').items.select("ID, Title, groupe_name").filter("Title eq '" + docTitle.toString() + "'").getAll();


    await Promise.all(allNotif.map(async (notif) => {

      html += `
          <tr>
          <td class="text-left">${notif.ID}</td>


          <td class="text-left">${notif.groupe_name}</td>
          
          <td class="text-left">

         <a href="#"  title="delete_notif" id="${notif.ID}_view_doc_notif" class="btn_delete_notif" style="padding-left: inherit;">
         <i class="fa-solid fa-trash" style="font-size: x-large;"></i>

     
         </a>

          </td>
        
         `;

    }))
      .then(() => {

        html += `</tbody>
          </table>`;
        listContainerDocNotif.innerHTML += html;
      });

    var table = $("#tbl_doc_notif").DataTable(

      {
        columnDefs: [
          {
            targets: [0],
            visible: false,
            target: 0
          }]

      }
    );


    $('#tbl_doc_notif tbody').on('click', '.btn_delete_notif', async (event) => {
      var data = table.row($(event.currentTarget).parents('tr')).data();
      await this._delete(data[0], "ToNotify", "Notification");
      window.location.reload();
    });

  }

  private async _delete(id: string, listName: string, dialog: string) {
    try {
      var res = await sp.web.lists.getByTitle(listName).items.getById(parseInt(id)).delete()
        .then(() => {
          alert(`${dialog} supprimé avec succès.`);
        })
        .then(() => {
          window.location.reload();
        });
    }
    catch (err) {
      alert(err.message);
    }

  }

  public async getSiteUsers() {

    var drp_users = document.getElementById("select_users");
    drp_users.innerHTML = "";


    const users1: any = await sp.web.siteUsers();

    users = users1;
    //console.log('SITEUSERSSSS ++++>', users1);

    users.forEach((result: ISiteUserInfo) => {

      if (result.UserPrincipalName != null) {

        //   console.log("USER", result.Id, result.Email);
        // if(result.IsFolder == "TRUE"){
        // console.log("SELECT_FOLDERS", result.Title);
        var opt = document.createElement('option');
        opt.appendChild(document.createTextNode(result.Email));
        opt.value = result.Email;
        drp_users.appendChild(opt);
        // }
      }

    });

  }

  public async getUserPrincipalId(graphClient: MSGraphClient, userId: string): Promise<string> {
    try {
      const user = await graphClient.api(`/users/${userId}`)
        .version('v1.0')
        .select('id,displayName,userPrincipalName,mail,userType,userPrincipalId')
        .get();

      return user.userPrincipalId;
    } catch (error) {
      console.error(`Error retrieving user ${userId} from Microsoft Graph API:`, error);
      throw error;
    }
  }

  public async getSiteGroups() {

    // var drp_users = document.getElementById("select_groups");
    var drp_users = document.getElementById("group") as HTMLSelectElement;

    if (!drp_users) {
      console.error("Dropdown element not found");
      return;
    }



    try {

      // const groups1: any = await sp.web.siteGroups();
      // const allGroups = await this.getGroupsByName(this.graphClient, "myGed");
      // console.log("groups", allGroups);

      //const { permissions } = await this.getSitePermissions('https://ncaircalin.sharepoint.com/sites/TestMyGed', "mgolapkhan.ext@aircalin.nc", "musharaf2897");
      const { permissions } = await this.getSitePermissions(this.context.pageContext.web.absoluteUrl, "mgolapkhan.ext@aircalin.nc", "musharaf2897");

      console.log("groups", permissions);

      for (const group of permissions) {

        var opt = document.createElement('option');
        // opt.appendChild(document.createTextNode(group.title));
        opt.value = group.title;

        opt.setAttribute('data-value', group.id);
        opt.dataset;

        drp_users.appendChild(opt);

      }


    } catch (error) {
      console.error("Error retrieving groups:", error);
    }

  }

  public async getSiteGroups_notif() {

    // var drp_users = document.getElementById("select_groups");
    var drp_users = document.getElementById("groups_notif") as HTMLSelectElement;

    if (!drp_users) {
      console.error("Dropdown element not found");
      return;
    }



    try {
      // const groups1: any = await sp.web.siteGroups();
      const allGroups = await this.getGroupsByName(this.graphClient, "myGed");
      // console.log("groups", allGroups);

      // const { permissions } = await this.getSitePermissions('https://ncaircalin.sharepoint.com/sites/TestMyGed', "mgolapkhan.ext@aircalin.nc", "musharaf2897");

      // console.log("groups", permissions);

      for (const group of allGroups) {

        var opt = document.createElement('option');
        // opt.appendChild(document.createTextNode(group.title));
        opt.value = group.displayName;

        opt.setAttribute('data-value', group.mail);
        opt.dataset;
        drp_users.appendChild(opt);

      }


    } catch (error) {
      console.error("Error retrieving groups:", error);
    }

  }

  public async getSitePermissions(siteUrl, username, password) {
    try {
      const credentials = btoa(`${username}:${password}`);
      const url = `${siteUrl}/_api/Web/RoleAssignments?$expand=Member,RoleDefinitionBindings`;

      const response = await fetch(url, {
        headers: {
          Accept: 'application/json;odata=verbose',
          Authorization: `Basic ${credentials}`,
        },
      });

      if (!response.ok) {
        throw new Error(`Network response was not ok: ${response.status}`);
      }

      const data = await response.json();
      const permissions = [];

      for (const entry of data.d.results) {
        const userOrGroup = entry.Member;

        if (userOrGroup.PrincipalType === 4) {
          // User or domain group
          const principalId = userOrGroup.Id;
          const title = userOrGroup.Title;
          const roleName = entry.RoleDefinitionBindings.results[0].Name;
          const email = userOrGroup.Email;

          // Add member to permissions
          permissions.push({ type: 'member', id: principalId, role: roleName, title: title, email: email });
        }
      }

      console.log("All Site Permissions:", permissions);

      return { permissions };
    } catch (err) {
      console.error(err);
      return { permissions: [] };
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


