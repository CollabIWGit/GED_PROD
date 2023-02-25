import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import $ from 'jquery';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DocDetailsWebPart.module.scss';
import * as strings from 'DocDetailsWebPartStrings';
import { sp, List, IItemAddResult, UserCustomActionScope, Items, Item, ITerm, ISiteGroup, ISiteGroupInfo } from "@pnp/sp/presets/all";
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



require('../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css');

SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.25/css/jquery.dataTables.min.css');

SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/jstree/3.2.1/themes/default/style.min.css');

SPComponentLoader.loadScript('https://ajax.googleapis.com/ajax/libs/jquery/3.6.3/jquery.min.js');

SPComponentLoader.loadScript("https://cdnjs.cloudflare.com/ajax/libs/jquery/1.12.1/jquery.min.js");
SPComponentLoader.loadScript('https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js');

// SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jstree/3.2.1/jstree.min.js');

require('./../../common/css/doctabs.css');
require('./../../common/css/minBootstrap.css');

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



export default class DocDetailsWebPart extends BaseClientSideWebPart<IDocDetailsWebPartProps> {

  private graphClient: MSGraphClient;
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';



  protected onInit(): Promise<void> {
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

  private getIfFolder() {
    var queryParms = new URLSearchParams(document.location.search.substring(1));
    var myParm = queryParms.get("folder");
    if (myParm) {
      return myParm.trim();
    }
  }

  private async _renderList() {
  }

  private async getParentID(id: any) {

    var parentID = null;
    var folderID = null;
    var parent_title = "";


    //var parentIDArray = [] ;

    await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + id + "'").get().then((results) => {

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

      await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID, Title").filter("FolderID eq '" + parentID + "'").get().then((results) => {

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


  public render(): void {




    //  SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css');
    // SPComponentLoader.loadScript('https://ajax.googleapis.com/ajax/libs/jquery/3.6.3/jquery.min.js');

    this.domElement.innerHTML = `

    <div class="wrapper d-flex align-items-stretch">
    
        <div id="jstree_demo_div"></div>
    
        <div class="jumbotron">
            <div class="row">
                <div class="col-md-7 top-buffer">
                    <h2 id="h2_doc_title">
                    </h2>
                </div>
                <div class="text-right inline" id="view_doc">
                    <h2>
                        <a href="#" target="_blank" id="open_doc" title="View document"> <i class="fa-regular fa-eye"
                                title="voir"></i></a>
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
                                        <label for="input_number">Nom du fichier</label>
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
    
                <div id="splistDocVersions" style="box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%); padding: 1em;">
    
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
                                        <input type="text" class="form-control" id="group_name" list='groups' />
    
                                        <datalist id="group">
                                            <select id="select_group"></select>
                                        </datalist>
                                    </div>
    
                                </div>
    
                                <div class="col-lg-4">
                                    <div class="form-group">
                                        <label for="permissions_group">Type</label>
                                        <select class='form-control' name="permissions" id="permissions_group">
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
    
                <div id="splistDocAccessRights" style="box-shadow: 0 4px 8px 0 rgb(0 0 0 / 20%), 0 6px 20px 0 rgb(0 0 0 / 19%); padding: 1em;">
    
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
                                            id="add_user">Ajouter</button>
                                    </div>
                                </div>
                            </div>
    
    
                            <div class="row">
                                <div class="col-lg-6">
    
                                    <div class="form-group">
                                        <label for="group_name_notif">Ajouter une notification de groupe :</label>
                                        <input type="text" class="form-control" id="group_name_notif" list='groups' />
    
                                        <datalist id="group">
                                            <select id="select_group"></select>
                                        </datalist>
                                    </div>
    
                                </div>
    
    
                                <div class="col-lg-3" style="padding-top: 1.7em;">
                                    <div class="form-group">
                                        <button type="button" class="btn btn-primary add_notif_group mb-2"
                                            id="add_group">Ajouter</button>
                                    </div>
                                </div>
                            </div>
    
                        </div>
    
    
    
    
                    </div>
                </div>
    
    
            </div>
    
            <div id="audit" class="tab-pane fade">
                <h3>Piste d'audit</h3>
    
    
    
            </div>
    
        </div>
    
    </div>
    </div>
    </div>
    `;

    var title = this.getDocTitle();
    var docId = this.getDocId();

    //require('./DocDetailsWebPartJS');



    this.getParentID(this.getDocId());
    this._getDocDetails(parseInt(docId));
    this._getAllVersions(title);
    this._getAllAccess(docId);
    this.getSiteGroups();
    this.getSiteUsers();
    this.fileUpload();


    // $('#jstree_demo_div').jstree(
    //   {
    //     "core": {
    //       "animation": 0,
    //       "check_callback": true,
    //       "themes": { "stripes": true },
    //       'data': {
    //         'url': function (node) {
    //           // Set the SharePoint site URL and list name
    //           var siteUrl = "https://frcidevtest.sharepoint.com/sites/myGed";
    //           var listName = "Documents";

    //           // Set the REST API URL based on the node ID
    //           var apiUrl = siteUrl + "/_api/web/lists/getbytitle('" + listName + "')/items";
    //           if (node.id !== '') {
    //             apiUrl += "?$filter=ParentID eq " + node.id; // replace ParentID with your own column name
    //           }

    //           console.log("RESPONSE", apiUrl);

    //           return apiUrl;
    //         },
    //         'headers': {
    //           'Accept': 'application/json;odata=verbose'
    //         },
    //         'data': function (node) {
    //           // Set any additional request parameters here
    //           return {};
    //         },
    //         'dataType': "json",
    //         'contentType': "application/json; charset=utf-8",
    //         'method': "GET",
    //         'processData': false,
    //         'success': function (data) {
    //           // Map the SharePoint list data to the expected jstree format
    //           var treeData = $.map(data.d.results, function (item) {
    //             return {
    //               'id': item.FolderID,
    //               'parent': item.ParentID, // replace ParentID with your own column name
    //               'text': item.Title // replace Title with your own column name
    //             };
    //           });
    //           console.log("TREE DATA", treeData);
    //           return treeData;
    //         }
    //       }
    //     },
    //     "types": {
    //       // Define your node types here
    //     },
    //     "plugins": [
    //       "contextmenu", "dnd", "search",
    //       "state", "types", "wholerow"
    //     ]
    //   }



    // );




    // this.createPath();

    $("#myTab a").click((e) => {
      e.preventDefault();
      (<any>$(this)).tab("show");

      table.columns.adjust().draw();

    });

    //update document
    $("#update_details_doc").click((e) => {
      this.updateDocument(docId, title);

    });

    $("#add_user").click((e) => {
      this.add_permission();
    });









  }

  private async updateDocument(folderId: string, title: string) {



    let user_current = await sp.web.currentUser();

    let text = $("#input_type_doc").val();
    const myArray = text.toString().split("_");
    let parentId = myArray[0];


    if (confirm(`Etes-vous sûr de vouloir mettre à jour les détails de ${title} ?`)) {

      try {

        const i = await await sp.web.lists.getByTitle('Documents').items.add({
          // const i = await await sp.web.lists.getByTitle('Documents').items.getById(parseInt(itemDoc.Id)).update({
          Title: $("#input_number").val(),
          description: $("#input_description").val(),
          keywords: $("#input_keywords").val(),
          doc_number: $("#input_number").val(),
          revision: $("#input_revision").val(),
          ParentID: parseInt(parentId),
          FolderID: folderId,
          filename: $("#input_filename").val(),
          IsFolder: "FALSE",
          owner: $("#created_by").val(),
          updatedBy: user_current.Title,
          createdDate: $("#creation_date").val(),
          updatedDate: new Date().toLocaleString()
        })
          .then(async (iar) => {

            var item = iar.data.ID;

            const list = sp.web.lists.getByTitle("Documents");

            await list.items.getById(iar.data.ID).attachmentFiles.add(filename_add, content_add);


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
    else {

    }






  }


  private fileUpload() {
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
        html += `<li><a href="https://frcidevtest.sharepoint.com/sites/myGed/SitePages/Home.aspx?folder=${item.parentId_doc}">${item.parentTitle_doc}</a></li>`;
      }
    });

    html += `</ul>`;

    listContainerDocPath.innerHTML += html;

  }

  private async _getDocDetails(id: number) {

    var urlFile_download = '';
    var titleFolder = '';
    var pdfNameDownload = '';

    const itemDoc: any = await sp.web.lists.getByTitle("Documents").items.getById(id)();

    await sp.web.lists.getByTitle("Documents")
      .items
      .getById(parseInt(itemDoc.Id))
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

    itemFolderId = itemDoc.FolderID;

    $("#input_type_doc").val(itemDoc.ParentID + "_" + titleFolder);
    //  $("#input_type_doc").val(itemDoc.ParentID);
    //input_type_doc
    $("#input_number").val(itemDoc.Title);
    $("#input_revision").val(itemDoc.revision);
    // $("#input_status").val(itemDoc.status);
    // $("#input_owner").val(itemDoc.owner);
    // $("#input_activeDate").val(itemDoc.active_date);
    // $("#input_filename").val(itemDoc.filename);
    // $("#input_author").val(itemDoc.author);
    // // $("#input_reviewDate").val(item1.);
    $("#input_keywords").val(itemDoc.keywords);
    $("#input_description").val(itemDoc.description);
    $("#created_by").val(itemDoc.owner);

    $("#updated_by").val(itemDoc.updateBy);
    $("#updated_time").val(itemDoc.updatedDate);

    // //   $("#creation_date").val(itemDoc.Created);
    $("#creation_date").val(itemDoc.createdDate);
    $("#h2_doc_title").text(itemDoc.Title);
    // $("#open_doc").attr("href", urlFile_download);

    $("#open_doc").click((e) => {
      window.open(`${urlFile_download}`, '_blank');

    });

  }

  private async add_permission() {

    //add permission user


    var ifFolder = "FALSE";
    var x = this.getDocId();
    var doc_title = "";
    var docID = "";


    const user: any = await sp.web.siteUsers.getByEmail($("#users_name").val().toString())();


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
        LoginName: user.Title
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
        <th class="text-left" >Actions</th>
      </tr>
    </thead>
    <tbody id="tbl_documents_versions_bdy">`;



      var response_doc_versions = null;
      var value1 = "FALSE";


      const all_documents_versions: any[] = await sp.web.lists.getByTitle('Documents').items
        .select("Id,ParentID,FolderID,Title,revision,IsFolder,description")
        .filter("Title eq '" + title + "' and IsFolder eq '" + value1 + "'")
        .get();


      response_doc_versions = all_documents_versions;

      if (response_doc_versions.length > 0) {
        $("#table_version_doc").css("display", "block");

        //  await Promise.all(contract.map(async (result) => {
        // await response_doc_versions.forEach(async (element_version) => {

        await Promise.all(all_documents_versions.map(async (element_version) => {

          var pdfName = '';
          var urlFile = '';


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


          html += `
          <tr>
          <td class="text-left">${element_version.Id}</td>

          <td class="text-left">${element_version.Title}</td>

          <td class="text-left"> 
          ${urlFile}          
          </td>

          <td class="text-left"> 
          ${element_version.revision}          
          </td>

          <td class="text-left"> 
          ${element_version.description}          
          </td>
          
          <td class="text-left">

         <a href="#"  title="Voir le document" id="${element_version.Id}_view_doc_version" class="btn_view_doc" style="padding-left: inherit;">
         <i class="fa-sharp fa-solid fa-eye" style="font-size: x-large;"></i>
         </a>

          </td>
        
         `;


        }))
          .then(() => {

            html += `</tbody>
          </table>`;
            listContainerDocVersions.innerHTML += html;
          });


        table = $("#tbl_doc_versions").DataTable(

          {

            order: [3, 'desc'],
            columnDefs: [
              {
                targets: [0, 2],
                visible: false,
              }]

          }
        );



        $('#tbl_doc_versions tbody').on('click', '.btn_view_doc', (event) => {
          var data = table.row($(event.currentTarget).parents('tr')).data();
          Navigation.navigate(data[2], true);
          //this.redirectToPage();
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

    const allPermissions: any[] = await sp.web.lists.getByTitle('AccessRights').items.select("ID,groupName,permission,FolderIDId, LoginName").filter("FolderIDId eq '" + parseInt(itemID) + "'").getAll();



    await Promise.all(allPermissions.map(async (perm) => {

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
          }]

      }
    );

    $('#tbl_doc_rights tbody').on('click', '.btn_delete_access', async (event) => {
      var data = table.row($(event.currentTarget).parents('tr')).data();
      await this._delete(data[0], "AccessRights", "Droits d'accès");
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

        console.log("USER", result.Id, result.Email);
        // if(result.IsFolder == "TRUE"){
        // console.log("SELECT_FOLDERS", result.Title);
        var opt = document.createElement('option');
        opt.appendChild(document.createTextNode(result.UserPrincipalName));
        opt.value = result.UserPrincipalName;
        drp_users.appendChild(opt);
        // }
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


