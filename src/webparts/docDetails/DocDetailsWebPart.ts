import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as $ from 'jquery';
import 'jquery-ui';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DocDetailsWebPart.module.scss';
import * as strings from 'DocDetailsWebPartStrings';
import { sp, List, IItemAddResult, UserCustomActionScope, Items, IItem, ISiteGroup, ISiteGroupInfo, Web, RoleDefinition, IRoleDefinition, ISiteUser } from "@pnp/sp/presets/all";
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
import * as moment from 'moment';
import 'downloadjs';
import { degrees, PDFDocument, radians, rgb, rotateDegrees, rotateRadians, StandardFonts, } from 'pdf-lib/cjs/api';
import download from 'downloadjs';
import { saveAs } from 'file-saver';
import { SiteGroups } from '@pnp/sp/site-groups';
import * as pdfjsLib from 'pdfjs-dist';
import Viewer from 'viewerjs';
import { PermissionKind } from "@pnp/sp/security";
import 'viewerjs/dist/viewer.css';





SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.25/css/jquery.dataTables.min.css');

// SPComponentLoader.loadScript("https://code.jquery.com/ui/1.12.1/jquery-ui.js");
// SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css");


SPComponentLoader.loadScript('');
SPComponentLoader.loadScript('');
SPComponentLoader.loadScript('');





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

  private async getParentID(id: any) {

    var parentID = null;
    var folderID = null;
    var parent_title = "";
    var value2 = "FALSE";
    var value1 = "TRUE";


    //var parentIDArray = [] ;

    await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + id + "' and IsFolder eq '" + value2 + "'").get().then((results) => {

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


    while (parentID != 1) {

      await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID, Title").filter("FolderID eq '" + parentID + "' and IsFolder eq '" + value1 + "'").get().then((results) => {

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





    this.createPath(parentTitle);


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
      const pdfDoc = await PDFDocument.load(existingPdfBytes);
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

  private async openPDFInIframe(url: string, filigraneText: string) {
    const pdfBytes = await this.generatePdfBytes(url, filigraneText);
    const pdfUrl = URL.createObjectURL(new Blob([pdfBytes], { type: 'application/pdf' }));

    const overlay = document.createElement('div');
    overlay.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background-color: rgba(0, 0, 0, 0.8);
      display: flex;
      justify-content: center;
      align-items: center;
      z-index: 9999;
    `;

    const iframe = document.createElement('iframe');
    iframe.src = pdfUrl;
    iframe.style.cssText = `
      border: none;
      width: 100%;
      height: 100%;
      max-width: 1000px;
      max-height: 90vh;
    `;
    // iframe.setAttribute('sandbox', 'allow-same-origin allow-popups allow-scripts');

    iframe.addEventListener('contextmenu', (event) => {
      event.preventDefault();
    });

    const closeButton = document.createElement('button');
    closeButton.innerText = 'Close';
    closeButton.style.cssText = `
      position: absolute;
      top: 20px;
      right: 20px;
      background-color: #fff;
      border: none;
      padding: 10px;
      cursor: pointer;
      font-size: 16px;
    `;

    closeButton.addEventListener('click', () => {
      document.body.removeChild(overlay);
    });

    overlay.appendChild(iframe);
    overlay.appendChild(closeButton);
    document.body.appendChild(overlay);
  }

  private async generatePdfBytes(fileUrl: string, filigraneText: string): Promise<Uint8Array> {
    try {
      const existingPdfBytes = await fetch(fileUrl).then(res => res.arrayBuffer());
      const pdfDoc = await PDFDocument.load(existingPdfBytes);

      const pages = await pdfDoc.getPages();

      const dateDownload = Date();

      for (const [i, page] of Object.entries(pages)) {
        const firstPage = pages[0];
        const { width, height } = firstPage.getSize();
        const helveticaFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
        const fontSize = 16;

        page.drawText(filigraneText + dateDownload, {
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

      return pdfBytes;
    } catch (e) {
      console.error('Failed to generate PDF bytes:', e);
      throw e;
    }
  }


  public async render(): Promise<void> {

    this.domElement.innerHTML = `


    <div class="wrapper d-flex align-items-stretch">
    
    <div id="loader" style="display: flex; align-items: center; justify-content: center; position: fixed; top: 0; left: 0; width: 100%; height: 100%; z-index: 9999; backdrop-filter: blur(5px);">
    <img src="https://ncaircalin.sharepoint.com/sites/TestMyGed/SiteAssets/images/loader.gif" alt="Loading..." />
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
    
                            <a href="#" id="delete_doc" role="button" title="delete document"> <i class="fa-solid fa-trash"
                                    title="Supprimer le document" style="
                                padding-left: 0.5em;"></i></a>
    
                            <a href="#" id="download_doc" role="button" title="Telecharger le document"> <i
                                    class="fa-solid fa-download" title="Telecharger le document" style="
                                padding-left: 0.5em;"></i></a>


                                <label class="switch" id="switch_fav">
                                <input type="checkbox" id="bookmark-switch" style="display: none;">
                                <i class="fa-regular fa-star star-icon" style="padding-left: 0.5em;"></i>
                              </label>
  
    
                        </h2>
                    </div>
                </div>
            </div>
    
            <div id="doc_path">
    
            </div>
    
            <ul class="nav nav-tabs" id="myTab">
                <li class="active"><a data-toggle="tab" href="#informations">Informations</a></li>
                <li><a data-toggle="tab" href="#versions">Toute Versions</a></li>
                <li><a data-toggle="tab" href="#access">Droits d'accès</a></li>
                <li><a data-toggle="tab" href="#notifications">Notifications</a></li>
                <li><a data-toggle="tab" href="#audit">Piste d'audit</a></li>
            </ul>
    
            <div class="tab-content">
    
                <div id="informations" class="tab-pane fade in active">
                    <h3>Informations</h3>
    
    
    
    
    
                    <div class="row" style="
      box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%);
      margin: 2em;
      padding: 2em;">
                        <div class="col-lg-12">
    
    
                            <div class="w3-container" id="form">
    
                                <legend>Détails</legend>
    
                                <div class="row">
                                    <div class="col-lg-6">
    
                                        <div class="form-group">
                                            <label for="input_number">Nom du document</label>
                                            <input type="text" id='input_number' class='form-control' disabled>
                                        </div>
    
                                    </div>
    
                                    <div class="col-lg-6">
                                        <div class="form-group">
                                            <label for="input_type_doc">Dossier</label>
                                            <input type="text" class="form-control" id="input_type_doc" list='folders'
                                                disabled />
    
                                            <datalist id="folders">
                                                <select id="select_folders"></select>
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
    
                                <div class="row">
                                    <div class="col-lg-6">
                                        <div class="form-group">
                                            <label for="input_number">Owner</label>
                                            <input type="text" id='created_by' class='form-control' disabled>
                                        </div>
                                    </div>
    
                                    <div class="col-lg-6">
                                        <div class="form-group">
                                            <label for="input_number">Date de création</label>
                                            <input type="text" id='creation_date' class='form-control' disabled>
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
                                            <label for="input_number">Date</label>
                                            <input id="input_reviewDate" name="myBrowser" class='form-control' type="text"
                                                readonly>
    
                                        </div>
                                    </div>
                                </div>
    
                                <div class="row">
                                    <div class="col-lg-6">
                                        <div class="form-group">
                                            <label for="input_number">Fichier</label>
                                            <input type="file" name="file" id="file_ammendment_update" class="form-control">
    
                                        </div>
                                    </div>
    
                                    <div class="col-lg-6">
                                        <div class="form-group">
                                            <label for="input_number">Nom du fichier local</label>
                                            <input type="text" id='input_filename' class='form-control' disabled />
    
                                        </div>
                                    </div>
                                </div>
    
                                <div class="row">
                                    <div class="col-lg-6">
                                        <div class="form-group">
                                            <label for="input_number">Updated by :</label>
                                            <input type="text" id='updated_by' class='form-control' disabled>
    
                                        </div>
                                    </div>
    
                                    <div class="col-lg-6">
                                        <div class="form-group">
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
                                        <button type="button" class="btn btn-primary" id='edit_cancel_doc'>Cancel</button>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
    
    
                </div>
    
                <div id="versions" class="tab-pane fade">
                    <h3>Toute Versions</h3>
    
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
                                                <option value="NONE">NONE</option>
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
                                                <option value="NONE">NONE</option>
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
                                            <input type="text" class="form-control" id="group_name_notif" list='group' />
    
                                            <datalist id="group">
                                                <select id="select_group"></select>
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
    
                    </div>
    
    
                </div>
    
            </div>
    
        </div>
    
    </div>
    

    `;

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



    require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');
    require('./../../common/jqueryui/jquery-ui.js');
    require('./../../common/css/doctabs.css');
    require('./../../common/css/minBootstrap.css');

    SPComponentLoader.loadScript('https://code.jquery.com/jquery-3.3.1.slim.min.js');
    SPComponentLoader.loadScript('https://cdn.jsdelivr.net/npm/popper.js@1.14.7/dist/umd/popper.min.js');
    SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.1.0/js/bootstrap.min.js');



    var title = this.getDocTitle();
    var docId = this.getDocId();

    //require('./DocDetailsWebPartJS');

    const loader = document.getElementById('loader');

    this.getParentID(this.getDocId());

    try {

      await this._getDocDetails(parseInt(docId)),
        // await this.checkPermission(),
        await this._getAllVersions(title),
        await this._getAllAccess(docId),
        await this._getAllAudit(docId),
        await this._getAllNotifications(docId),
        await this.getSiteGroups(),
        await this.getSiteUsers(),
        await this.load_folders(),
        this.fileUpload();

      $("#loader").css("display", "none");

    } catch (error) {
      // $("#loader").html(`Error: ${error.message}`);
      $("#loader").css("display", "none");

      console.log(error.message);
    }
    //update document


    //add_permission user
    $("#add_user").click(async (e) => {
      await this.add_permission($("#users_name").val().toString());
    });

    //add_permission_group
    $("#add_group").click(async (e) => {

      const stringGroupUsers: string[] = await this.getAllUsersInGroup($("#group_name").val());
      console.log("TESTER GROUP USERS", stringGroupUsers);
      await this.add_permission_group(stringGroupUsers);
    });

    //add group notif
    $("#add_group_notif").click(async (e) => {

      const stringGroupUsers: string[] = await this.getAllUsersInGroup($("#group_name_notif").val());
      console.log("TESTER GROUP USERS", stringGroupUsers);
      await this.add_notification_group(stringGroupUsers);
    });


    $("#add_user_notif").click((e) => {
      this.add_notification();
    });


    // var isBookmark = localStorage.getItem('bookmark') === 'true';
    // if (isBookmark) {
    //   $('#bookmark-switch').prop('checked', true);

    // }

    // Toggle the switch

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

  // private async checkUserPermissions( listTitle: string, itemId: number, userId: number): Promise<PermissionKind[]> {
  //   try {
  //     // Get the item with ID 'itemId' from the list with title 'listTitle' using PnP JS library.
  //     const item: any = await sp.web.lists.getByTitle(listTitle).items.getById(itemId).get();

  //     // Get the site user with ID 'userId' using PnP JS library.
  //     const user: any = await sp.web.siteUsers.getById(userId).get();

  //     // Get the user's effective permissions on the item using PnP JS library.
  //     const userPermissions: PermissionKind[] = await item.getEffectiveBasePermissions(user.LoginName);

  //     return userPermissions;
  //   } catch (error) {
  //     console.error(error);
  //     throw error;
  //   }
  // }

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

  private async getAllGroups(graphClient: MSGraphClient): Promise<any[]> {
    try {
      const groups = await graphClient.api('/groups')
        .version('v1.0')
        .get();

      return groups.value;
    } catch (error) {
      console.error(error);
      return [];
    }
  }

  private async updateDocument(folderId: string, title: string) {

    let user_current = await sp.web.currentUser();

    let text = $("#input_type_doc").val();
    const value1 = "TRUE";
    var folder = '';


    const all_folders = await sp.web.lists.getByTitle('Documents').items
      .select("ID,ParentID,FolderID,Title,IsFolder,description")
      .top(5000)
      .filter("Title eq '" + text + "' and IsFolder eq '" + value1 + "' ")
      .get();

    await Promise.all(all_folders.map(async (doc) => {

      folder = doc.FolderID;

    }));

    // const myArray = text.toString().split("_");
    // let parentId = myArray[0];

    if ($('#file_ammendment').val() == '') {

      alert("Veuillez télécharger le fichier avant de continuer.");

    }

    else {

      if (confirm(`Etes-vous sûr de vouloir mettre à jour les détails de ${title} ?`)) {

        try {
          const i = await sp.web.lists.getByTitle('Documents').items.add({
            Title: $("#input_number").val(),
            description: $("#input_description").val(),
            keywords: $("#input_keywords").val(),
            doc_number: $("#input_number").val(),
            revision: $("#input_revision").val(),
            ParentID: parseInt(folder),
            IsFolder: "FALSE",
            owner: $("#created_by").val(),
            updatedBy: user_current.Title,
            createdDate: $("#creation_date").val(),
            updatedDate: new Date().toLocaleString()

          })
            .then(async (iar) => {

              const list = sp.web.lists.getByTitle("Documents");

              let folderId_link = iar.data.ID;
              let title = iar.data.Title;

              await list.items.getById(iar.data.ID).attachmentFiles.add(filename_add, content_add);

              return { list, title, folderId_link };

            })
            .then(async ({ list, title, folderId_link }) => {

              await list.items.getById(folderId_link).update({
                FolderID: parseInt(folderId_link),
                filename: filename_add
              });

              return { title, folderId_link };

            })
            .then(async ({ title, folderId_link }) => {
              await this.createAudit($("#input_number").val(), folderId, user_current.Title, "Modification");
              return { title, folderId_link }
            })
            .then(({ title, folderId_link }) => {
              // window.location.href = `https://ncaircalin.sharepoint.com/sites/TestMyGed/SitePages/Document.aspx?document=${title}&documentId=${folderId_link}`;

              alert("Détails mis à jour avec succès");
              return { title, folderId_link };
            })
            .then(({ title, folderId_link }) => {
              //    location.reload(true)
              window.location.href = `https://ncaircalin.sharepoint.com/sites/TestMyGed/SitePages/Document.aspx?document=${title}&documentId=${folderId_link}`;

            });

        }
        catch (err) {
          alert(err.message);
        }


      }
      else {

      }
    }

  }

  private async checkPermission() {
    const groupTitle = [];
    let groups: any = await sp.web.currentUser.groups();

    console.log("PERMISSION", groups);

    await Promise.all(groups.map(async (perm) => {

      groupTitle.push(perm.Title);

    }));

    // if (groupTitle.includes("myGed Visitors")) {
    if (groupTitle.includes("Utilisateur MyGed")) {

      $("#update_details_doc, #edit_cancel_doc, #access, #notifications, #audit, #delete_doc, #download_doc").css("display", "none");

      $("#input_description, #input_keywords, #input_revision, #file_ammendment_update ").prop('disabled', true);

    }
    else if (groupTitle.includes("Référent (Read & Write)")) {

      $("#access").css('display', "none");
      // $("#update_details_doc, #edit_cancel_doc, #access, #notifications, #audit").css("display", "block");
    }

    else if (groupTitle.includes("Administrateur")) {

      $("#input_number, #input_type_doc").prop('disabled', false);
      // $("#update_details_doc, #edit_cancel_doc, #access, #notifications, #audit").css("display", "block");
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

  private async _getAllAudit(id: string) {

    var value1 = "FALSE";
    var itemID = "";

    var doc_detail: any = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + parseInt(id) + "' and IsFolder eq '" + value1 + "'").get();

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

    const allAudit: any[] = await sp.web.lists.getByTitle('Audit').items.select("ID, Title, Person, FolderID, Action, DateCreated").filter("FolderID eq '" + parseInt(id) + "'").getAll();


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

    var table = $("#tbl_doc_audit").DataTable();

  }

  private async add_permission(user_group: any) {

    //add permission user


    var ifFolder = "FALSE";
    var x = this.getDocId();
    var doc_title = "";
    var docID = "";


    const user: any = await sp.web.siteUsers.getByEmail(user_group)();



    const all_documents: any[] = await sp.web.lists.getByTitle('Documents').items
      .select("Id,ParentID,FolderID,Title,revision,IsFolder,description")
      .filter("FolderID eq '" + x + "' and IsFolder eq '" + ifFolder + "'")
      .get();

    await Promise.all(all_documents.map(async (doc) => {

      doc_title = doc.Title;
      docID = doc.Id;


    }));


    console.log("USERS FOR PERMISSION", users_Permission);


    try {
      // response_same_doc.forEach(async (x) => {

      await sp.web.lists.getByTitle("AccessRights").items.add({
        Title: doc_title.toString(),
        groupName: $("#users_name").val(),
        permission: $("#permissions_user option:selected").val(),
        FolderIDId: docID.toString(),
        PrincipleID: user.Id,
        LoginName: user.Title,
        groupTitle: $("#group_name").val()
      })
        .then(() => {
          alert("Autorisation ajoutée à ce document avec succès.");
        })
        .then(() => {
          window.location.reload();
        });

      // });

      // alert("Autorisation ajoutée à ce document avec succès.");
      // window.location.reload();
    }

    catch (e) {
      alert("Erreur: " + e.message);
    }


  }

  private async getUserPermissionLevelOnSharePointListItem(listName: string, itemId: number, userEmail: string): Promise<string> {
    try {
      const item = await sp.web.lists.getByTitle(listName).items.getById(itemId).get();
      const userPermission = await item.roleAssignments.get().filter(`Member/Email eq '${userEmail}'`).get();

      if (userPermission && userPermission.length > 0) {
        const permissionLevel = userPermission[0].roleDefinitionBindings[0].Name;
        return permissionLevel;
      }
    } catch (error) {
      console.error(error);
    }

    return 'No Access';
  }

  private async add_permission_group(group_name: string[]) {

    //add permission user

    var ifFolder = "FALSE";
    var x = this.getDocId();
    var doc_title = "";
    var docID = "";

    const all_documents: any[] = await sp.web.lists.getByTitle('Documents').items
      .select("Id,ParentID,FolderID,Title,revision,IsFolder,description")
      .filter("FolderID eq '" + x + "' and IsFolder eq '" + ifFolder + "'")
      .get();

    await Promise.all(all_documents.map(async (doc) => {
      doc_title = doc.Title;
      docID = doc.Id;
    }));

    console.log("USERS FOR PERMISSION", group_name);

    try {
      await Promise.all(group_name.map(async (email) => {
        const user: any = await sp.web.siteUsers.getByEmail(email)();
        await sp.web.lists.getByTitle("AccessRights").items.add({
          Title: doc_title.toString(),
          groupName: email,
          permission: $("#permissions_group option:selected").val(),
          FolderIDId: docID.toString(),
          PrincipleID: user.Id,
          LoginName: user.Title,
          groupTitle: $("#group_name").val()
        });
      }));

      alert("Authorization added successfully.");
      window.location.reload();
    }
    catch (e) {
      alert("Error: " + e.message);
    }
  }

  private async add_notification_group(group_name: string[]) {

    //add permission group

    var ifFolder = "FALSE";
    var x = this.getDocId();
    var doc_title = "";
    var docID = "";
    var revisionDate = "";
    var description = "";
    var revision = "";


    const all_documents: any[] = await sp.web.lists.getByTitle('Documents').items
      .select("Id,ParentID,FolderID,Title,revision,IsFolder,description")
      .filter("FolderID eq '" + x + "' and IsFolder eq '" + ifFolder + "'")
      .get();

    await Promise.all(all_documents.map(async (doc) => {
      doc_title = doc.Title;
      docID = doc.Id;
      revisionDate = doc.revisionDate;
      description = doc.description;
      revision = doc.revision;

    }));

    console.log("USERS FOR PERMISSION", group_name);

    try {
      await Promise.all(group_name.map(async (email) => {
        const user: any = await sp.web.siteUsers.getByEmail(email)();
        await sp.web.lists.getByTitle("Notifications").items.add({
          Title: doc_title.toString(),
          group_person: email,
          IsFolder: "FALSE",
          revisionDate: revisionDate,
          toNotify: "YES",
          description: description,
          FolderID: x.toString(),
          webLink: `https://ncaircalin.sharepoint.com/sites/TestMyGed/SitePages/Document.aspx?document=${doc_title}&documentId=${x}`,
          LoginName: user.Title,
          revision: revision
        })
      }));

      alert("Notification ajoutée à ce document avec succès.");
      window.location.reload();
    }
    catch (e) {
      alert("Error: " + e.message);
    }
  }

  private async load_folders() {
    const value1 = "TRUE";
    const drp_folders = document.getElementById("select_folders");

    if (!drp_folders) {
      console.error("Dropdown element not found");
      return;
    }

    const all_folders = await sp.web.lists.getByTitle('Documents').items
      .select("ID,ParentID,FolderID,Title,IsFolder,description")
      .top(5000)
      .filter("IsFolder eq '" + value1 + "'")
      .get();

    folders = all_folders;

    folders.forEach((result: any) => {
      const opt = document.createElement('option');
      opt.appendChild(document.createTextNode(result.Title));
      opt.value = result.Title;
      drp_folders.appendChild(opt);
    });
  }

  private fileUpload() {

  }

  private createPath(listDoc: any) {

    const listContainerDocPath: Element = this.domElement.querySelector('#doc_path');
    let html: string = `<ul class="breadcrumb" id="breadcrumb">`;


    // for (var i in parentTitle) {
    //   if (parentTitle[i] !== undefined) {
    //     html += `<li><a href="#">${parentTitle[i]}</a></li>`;
    //     console.log("OLLL", parentTitle[i]);
    //   }
    // }

    listDoc.forEach((item) => {

      if (item.parentTitle_doc !== undefined) {
        html += `<li><a href="${this.context.pageContext.web.absoluteUrl}/SitePages/Home.aspx?folder=${item.parentId_doc}">${item.parentTitle_doc}</a></li>`;
      }
    });

    html += `</ul>`;

    listContainerDocPath.innerHTML += html;

  }

  // private async showPDF(url: any, filigraneText: any) {
  //   const pdfBytes = await this.generatePdfBytes(url, filigraneText);
  //   const pdfUrl = URL.createObjectURL(new Blob([pdfBytes], { type: 'application/pdf' }));

  //   // Create a container element to hold the viewer
  //   const container = document.createElement('div');
  //   container.setAttribute('style', 'position:fixed;top:0;left:0;width:100%;height:100%;z-index:1000;background-color:#fff;');
  //   document.body.appendChild(container);

  //   // Create an iframe element to display the PDF file
  //   const iframe = document.createElement('iframe');
  //   iframe.setAttribute('src', 'viewer.html');
  //   iframe.setAttribute('style', 'width:100%;height:100%;border:none;');
  //   iframe.setAttribute('scrolling', 'no');
  //   iframe.setAttribute('allowfullscreen', 'true');

  //   // Add the iframe to the container element
  //   container.appendChild(iframe);

  //   // Wait for the iframe to load
  //   iframe.addEventListener('load', () => {
  //     // Get a reference to the iframe's content window
  //     const iframeWindow = iframe.contentWindow as Window;

  //     // Load Viewer.js in the iframe
  //     const script = iframeWindow.document.createElement('script');
  //     script.setAttribute('src', 'viewer.js');
  //     script.addEventListener('load', () => {
  //       // Initialize viewer.js with the PDF file
  //       const viewer = new iframeWindow.Viewer({
  //         inline: true,
  //         button: false,
  //         toolbar: false,
  //         navbar: false,
  //         fullscreen: true,
  //         url: pdfUrl
  //       });
  //       viewer.show();
  //     });

  //     // Add the Viewer.js script to the iframe's document
  //     iframeWindow.document.body.appendChild(script);
  //   });
  // }

  private async showPDF(url: string, filigraneText: string) {

    pdfjsLib.GlobalWorkerOptions.workerSrc = require('../../../node_modules/pdfjs-dist/build/pdf.worker.js');


    // Load the PDF document
    const pdf = await pdfjsLib.getDocument(url).promise;

    // Get the total number of pages in the PDF document
    const numPages = pdf.numPages;

    // Create a container element to hold the viewer and close button
    const container = document.createElement('div');
    container.setAttribute('style', 'position:fixed;top:0;left:0;width:100%;height:100%;z-index:1000;background-color:#fff;');
    document.body.appendChild(container);

    // Create a close button element
    const closeButton = document.createElement('button');
    closeButton.innerHTML = 'Close';
    closeButton.addEventListener('click', () => {
      document.body.removeChild(container);
    });
    container.appendChild(closeButton);

    // Create a div element to hold the pages of the PDF document
    const pagesContainer = document.createElement('div');
    pagesContainer.setAttribute('style', 'width:100%;height:calc(100% - 30px);overflow:auto;');
    container.appendChild(pagesContainer);

    // Loop through each page of the PDF document and add it to the pages container
    for (let i = 1; i <= numPages; i++) {
      // Load the page
      const page = await pdf.getPage(i);

      // Get the page viewport
      const viewport = page.getViewport({ scale: 1 });

      // Create a canvas element to display the page
      const canvas = document.createElement('canvas');
      canvas.setAttribute('style', 'display:block;margin:10px auto;border:1px solid #ccc;');
      canvas.height = viewport.height;
      canvas.width = viewport.width;

      // Add the canvas element to the pages container
      pagesContainer.appendChild(canvas);

      // Render the page on the canvas
      await page.render({
        canvasContext: canvas.getContext('2d'),
        viewport
      }).promise;

      // Add the filigrane text to the canvas
      const context = canvas.getContext('2d');
      context.fillStyle = 'rgba(255, 255, 255, 0.5)';
      context.fillRect(0, 0, canvas.width, canvas.height);
      context.font = '24px Arial';
      context.fillStyle = 'red';
      context.fillText(filigraneText, 20, 40);
    }
  }

  private async _getDocDetails(id: number) {

    // var externalUrl = '';
    // var url = '';
    var urlFile_download = '';
    var pdfNameDownload = '';

    await this.checkPermission();
    var x = await this.getAllGroups(this.graphClient);

    console.log("ALL AD GROUPS", x);


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
    const itemDoc: any[] = await sp.web.lists.getByTitle('Documents').items
      .select("Id,ParentID,FolderID,Title,revision,IsFolder, description, revisionDate, keywords, owner, updatedBy, updatedDate, createdDate, attachmentUrl, IsFiligrane, IsDownloadable")
      .filter("FolderID eq '" + id + "' and IsFolder eq 'FALSE'")
      .get();

    // Get the folder details
    const itemFolder: any[] = await sp.web.lists.getByTitle('Documents').items
      .select("Id,ParentID,FolderID,Title")
      .filter("FolderID eq '" + parseInt(itemDoc[0].ParentID) + "' and IsFolder eq 'TRUE'")
      .get();



    await sp.web.lists.getByTitle("Documents")
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



    // Set the document details in the UI
    $("#input_type_doc").val(itemFolder[0].Title);
    $("#input_number").val(itemDoc[0].Title);
    $("#input_revision").val(itemDoc[0].revision);
    $("#input_keywords").val(itemDoc[0].keywords);
    $("#input_description").val(itemDoc[0].description);
    $("#created_by").val(itemDoc[0].owner);
    $("#updated_by").val(itemDoc[0].updatedBy);
    $("#updated_time").val(itemDoc[0].updatedDate);
    $("#creation_date").val(itemDoc[0].createdDate);
    $("#h2_doc_title").text(itemDoc[0].Title);

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

    var xx = await this.getPermissionLevel(this.graphClient, "TestMyGed", "Documents", itemDoc[0].Id);
    console.log("GRAPH PERMISSION", xx);

    // Open the attachment in a new tab

    let user_current = await sp.web.currentUser();
    $("#open_doc").click(async (e) => {

      if (itemDoc[0].IsFiligrane === "NO") {
        window.open(`${url}`, '_blank');
      }

      else {

        // else if (itemDoc[0].IsFiligrane === "YES") {

        await this.openPDFInIframe(url, 'UNCONTROLLED COPY - Downloaded on: ');

        //  await this.showPDF(url, 'UNCONTROLLED COPY - Downloaded on: ');
      }

      await this.createAudit(itemDoc[0].Title, itemDoc[0].FolderID, user_current.Title, "Consultation");
    });


    //check user permission
    var userPermission = this.getUserPermissionLevelOnSharePointListItem("Documents", itemDoc[0].Id, user_current.Email);
    console.log("User Permission", userPermission);


    // Delete the document
    $("#delete_doc").click(async (e) => {
      if (confirm(`Are you sure you want to delete ${itemDoc[0].Title}?`)) {
        try {
          await sp.web.lists.getByTitle('Documents').items.getById(parseInt(itemDoc[0].Id)).recycle();
          alert("Document deleted successfully.");
          window.location.reload();
        } catch (err) {
          alert(err.message);
        }
      }
    });

    // Download the attachment
    $("#download_doc").click(async (e) => {
      const user = await sp.web.currentUser();
      await this.downloadDoc(url, pdfNameDownload, itemDoc[0].FolderID, 'UNCONTROLLED COPY - Downloaded on ');
    });


    $("#update_details_doc").click(async (e) => {
      const fileInput = document.getElementById("file_ammendment_update") as HTMLInputElement;

      if (($("#input_type_doc").val() === itemFolder[0].Title ||
        $("#input_number").val() === itemDoc[0].Title ||
        $("#input_keywords").val() === itemDoc[0].keywords ||
        $("#input_description").val() === itemDoc[0].description) && $("#input_revision").val() !== itemDoc[0].revision && fileInput.value) {
        // The fields are unchanged, so update the document metadata
        await this.updateDocument(itemDoc[0].Id, itemDoc[0].Title);

        alert("You have created a new version of : " + itemDoc[0].Title);
      } else {
        // The fields are changed, so get the folder details and update the document metadata
        const folder = await this.getFolder();
        await this.updateDocMetadata(itemDoc[0].Id, folder.ID, itemDoc[0].FolderID, user_current.Title);
        alert("You have modified some metadata of : " + itemDoc[0].Title);
      }

    });


    const bookmarkSwitch = document.getElementById("bookmark-switch") as HTMLInputElement;
    bookmarkSwitch.addEventListener("change", async () => {
      await this.handleBookmarkSwitchChange.call(this, bookmarkSwitch.checked, itemDoc[0].FolderID, itemDoc[0].Title);
    });

    const user: ISiteUserInfo = await sp.web.currentUser();

    // await this.checkUserPermissions( "Documents", itemDoc[0].ID , user.Id);


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
        alert("You have set this document as favorite.");
        window.location.reload();
      } else {
        await this.removeBookmark(doc_id);
        alert("You have removed this document from favorites.");
        window.location.reload();
      }
    } catch (error) {
      console.error("Failed to update bookmark:", error);
    }
  }

  private async getFolder() {
    let folder = { ID: '', Title: '' };
    const items = await sp.web.lists.getByTitle('Documents').items
      .select("Id,ParentID,FolderID,Title")
      .filter(`Title eq '${$("#input_type_doc").val().toString()}' and IsFolder eq 'TRUE'`)
      .get();
    if (items.length > 0) {
      folder.ID = items[0].Id;
      folder.Title = items[0].Title;
    }
    return folder;
  }

  private async updateDocMetadata(id: any, folder: any, folderID: any, userTitle: any) {

    try {
      const list = sp.web.lists.getByTitle("Documents");

      const i = await list.items.getById(id).update({
        Title: $("#input_number").val(),
        description: $("#input_description").val(),
        keywords: $("#input_keywords").val(),
        doc_number: $("#input_number").val(),
        ParentID: Number(folder),

      }).then(async () => {
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

  private async add_notification() {

    //add permission user


    var ifFolder = "FALSE";
    var x = this.getDocId();
    var doc_title = "";
    var docID = "";
    var revisionDate = "";
    var description = "";
    var link = "";


    const user: any = await sp.web.siteUsers.getByEmail($("#users_name_notif").val().toString())();


    const all_documents: any[] = await sp.web.lists.getByTitle('Documents').items
      .select("Id,ParentID,FolderID,Title,revision,IsFolder, description, revisionDate")
      .filter("FolderID eq '" + x + "' and IsFolder eq '" + ifFolder + "'")
      .get();



    await Promise.all(all_documents.map(async (doc) => {

      doc_title = doc.Title;
      docID = doc.Id;
      revisionDate = doc.revisionDate;
      description = doc.description;
      link = `https://ncaircalin.sharepoint.com/sites/TestMyGed/SitePages/Document.aspx?document=OWNERTEST&documentId=${doc.FolderID}`;

    }));


    console.log("USERS FOR PERMISSION", users_Permission);
    console.log("LIIINK", link);


    try {


      await sp.web.lists.getByTitle("Notifications").items.add({
        Title: doc_title.toString(),
        group_person: $("#users_name_notif").val(),
        IsFolder: "FALSE",
        revisionDate: revisionDate,
        toNotify: "YES",
        description: description,
        FolderID: x.toString(),
        webLink: `https://ncaircalin.sharepoint.com/sites/TestMyGed/SitePages/Document.aspx?document=${doc_title}&documentId=${x}`,
        LoginName: user.Title,
      })
        .then(() => {
          alert("Notification ajoutée à ce document avec succès.");
        })
        .then(() => {
          window.location.reload();
        });


    }

    catch (e) {
      alert("Erreur: " + e.message);
    }

  }

  private async _getAllVersions(title: string) {

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


      const all_documents_versions: any[] = await sp.web.lists.getByTitle('Documents').items
        .select("Id,ParentID,FolderID,Title,revision,IsFolder,description, attachmentUrl, IsFiligrane, IsDownloadable")
        .filter("Title eq '" + title + "' and IsFolder eq '" + value1 + "'")
        .get();



      // First, sort the array by revision in descending order
      // all_documents_versions.sort((a, b) => b.revision - a.revision);

      // Then, find the highest revision
      const highestRevision = all_documents_versions[0].revision;

      // Finally, filter the array and remove the items with the highest revision
      //  const filtered_documents_versions_1 = all_documents_versions.filter((document) => document.revision < highestRevision);

      const filtered_documents_versions_2 = all_documents_versions.filter((document) => document.revision !== null);


      if (filtered_documents_versions_2.length > 0) {
        $("#table_version_doc").css("display", "block");

        //  await Promise.all(contract.map(async (result) => {
        // await response_doc_versions.forEach(async (element_version) => {

        await Promise.all(filtered_documents_versions_2.map(async (element_version) => {

          var pdfName = '';
          var urlFile = '';
          var url = '';
          var attachmentUrl = element_version.attachmentUrl;


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

          if (data[7] === "NO") {
            alert("FILIGRANE = NO");
            window.open(`${data[2]}`, '_blank');
          }

          else if (data[7] === "YES") {

            alert("FILIGRANE = YES");
            //   await this.openPDFInBrowser(url, 'UNCONTROLLED COPY - Downloaded on: ');
            await this.openPDFInIframe(data[2], 'UNCONTROLLED COPY - Downloaded on: ');
          }
          //  window.open(`${data[2]}`, '_blank');
        });


        $('#tbl_doc_versions tbody').on('click', '.btn_download_doc', async (event) => {
          var data = table.row($(event.currentTarget).parents('tr')).data();
          await this.downloadDoc(data[2], data[5], data[6], 'ARCHIVED COPY - Downloaded on ');
        });

      }

    }

  }

  private async _getAllAccess(id: string) {

    var value1 = "FALSE";
    var itemID = "";

    var doc_detail: any = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + parseInt(id) + "' and IsFolder eq '" + value1 + "'").get();

    await Promise.all(doc_detail.map(async (perm) => {

      itemID = perm.Id;

    }));

    const listContainerDocRights: Element = this.domElement.querySelector('#splistDocAccessRights');

    let html: string = `<table id='tbl_doc_rights' class='table table-striped' style="width: 100%;font-size: initial;" >`;

    html += `<thead>
    <tr>
      <th class="text-left">ID</th>
      <th class="text-left">Nom</th>
      <th class="text-left" >Droits d'accès</th>

      <th class="text-left" >Actions</th>
    </tr>
  </thead>
  <tbody id="tbl_documents_access_bdy">`;

    const allPermissions: any[] = await sp.web.lists.getByTitle('AccessRights').items.select("ID,groupName,permission,FolderIDId,LoginName, groupTitle, Created").filter("FolderIDId eq '" + Number(itemID) + "'").getAll();

    const permissionsByGroup = allPermissions.reduce((groups, permission) => {
      const groupName = permission.groupTitle;
      if (!groups[groupName] || permission.Created > groups[groupName].Created) {
        groups[groupName] = permission;
      }
      return groups;
    }, {});

    const mostRecentPermissions: any = Object.values(permissionsByGroup);


    await Promise.all(mostRecentPermissions.map(async (perm) => {

      html += `
          <tr>
          <td class="text-left">${perm.Id}</td>

          <td class="text-left">${perm.LoginName}</td>
          <td class="text-left">${perm.permission}</td>
          
          <td class="text-left">

         <a href="#"  title="delete_perm" id="${perm.Id}_view_doc_version" class="btn_delete_access" style="padding-left: inherit;">
         <i class="fa-solid fa-trash" style="font-size: x-large;"></i>

     
         </a>

          </td>

         `;


    }))
      .then(() => {

        html += `</tbody>
          </table>`;
        listContainerDocRights.innerHTML += html;
      });

    var table = $("#tbl_doc_rights").DataTable(

      {
        columnDefs: [
          {
            targets: [0],
            visible: false,
            target: 0
          }]

      }
    );

    $('#tbl_doc_rights tbody').on('click', '.btn_delete_access', async (event) => {
      var data = table.row($(event.currentTarget).parents('tr')).data();
      await this._delete(data[0], "AccessRights", "Droits d'accès");
      window.location.reload();
    });



  }

  private async _getAllNotifications(id: string) {

    var value1 = "FALSE";
    var itemID = "";
    var folderID = "";

    var doc_detail: any = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + parseInt(id) + "' and IsFolder eq '" + value1 + "'").get();

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

    const allNotif: any[] = await sp.web.lists.getByTitle('Notifications').items.select("ID, Title, group_person, revisionDate, toNotify, webLink, description, FolderID").filter("FolderID eq '" + folderID.toString() + "'").getAll();

    await Promise.all(allNotif.map(async (notif) => {

      html += `
          <tr>
          <td class="text-left">${notif.Id}</td>


          <td class="text-left">${notif.group_person}</td>
          
          <td class="text-left">

         <a href="#"  title="delete_notif" id="${notif.Id}_view_doc_notif" class="btn_delete_notif" style="padding-left: inherit;">
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
      await this._delete(data[0], "Notifications", "Notification");
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
    // var drp_users2 = document.getElementById("select_users2");


    if (drp_users == null) {
      console.error("select_users element not found");
      return;
    }

    // drp_users.innerHTML = "";
    // drp_users2.innerHTML = "";

    const users1: any = await sp.web.siteUsers();

    users = users1;

    users.forEach((result: ISiteUserInfo) => {
      if (result.UserPrincipalName != null) {
        var opt = document.createElement('option');
        opt.appendChild(document.createTextNode(result.Email));
        opt.value = result.Email;
        drp_users.appendChild(opt);
      }
    });
  }

  public async getSiteGroups() {

    var drp_users = document.getElementById("select_groups");


    const groups1: any = await sp.web.siteGroups();

    groups = groups1;
    //console.log('SITEUSERSSSS ++++>', users1);

    groups.forEach((result: ISiteGroupInfo) => {

      if (result.Title != null) {
        //  console.log("USER", result.Email);
        // if(result.IsFolder == "TRUE"){
        // console.log("SELECT_FOLDERS", result.Title);
        var opt = document.createElement('option');
        opt.appendChild(document.createTextNode(result.LoginName));
        opt.value = result.LoginName;
        drp_users.appendChild(opt);
        // }
      }

    });

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


