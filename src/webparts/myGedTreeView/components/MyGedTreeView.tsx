import styles from './MyGedTreeView.module.scss';
import { MSGraphClient } from '@microsoft/sp-http';
import { IMyGedTreeViewProps, IMyGedTreeViewState } from './IMyGedTreeView';
import { TreeView, ITreeItem, TreeViewSelectionMode, TreeItemActionsDisplayMode } from "@pnp/spfx-controls-react/lib/TreeView";
import 'bootstrap/dist/css/bootstrap.min.css';
import $, { event } from 'jquery';
import Popper from 'popper.js';
import 'bootstrap/dist/js/bootstrap.bundle.min';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { sp, List, IItemAddResult, UserCustomActionScope, Items, Item, ITerm, ISiteGroup, ISiteGroupInfo } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { getIconClassName, Label, rgb2hex } from 'office-ui-fabric-react';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faFolder, faFolderOpen, faFileWord, faEdit, faTrashCan, faBell, faEye } from '@fortawesome/free-regular-svg-icons'
import { faFile, faLock, faFolderPlus, faDownload, faMagnifyingGlass, faDeleteLeft, faCircleInfo } from '@fortawesome/free-solid-svg-icons'
import { icon, IconName, IconProp, parse } from '@fortawesome/fontawesome-svg-core';
import React, { useEffect, useLayoutEffect, useState } from 'react';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import { IAttachmentInfo } from "@pnp/sp/attachments";
import "@pnp/sp/attachments";
import { IItem } from "@pnp/sp/items/types";
// import Form from 'react-bootstrap-form';
import * as sharepointConfig from './../../../common/utils/sharepoint-config.json';
import "@pnp/sp/site-groups/web";
import { SPComponentLoader } from '@microsoft/sp-loader';
import 'datatables.net';
SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.25/css/jquery.dataTables.min.css');
import Form from 'react-bootstrap/Form';
import { degrees, PDFDocument, radians, rgb, rotateDegrees, rotateRadians, StandardFonts, } from 'pdf-lib/cjs/api';
import 'downloadjs';
import * as download from 'downloadjs';
import "@pnp/sp/security/web";
import "@pnp/sp/site-users/web";
import { IList } from "@pnp/sp/lists";
import { PermissionKind } from "@pnp/sp/security";
// import Button from 'react-bootstrap/Button';
// import Modal from 'react-bootstrap/Modal';
import * as moment from 'moment';
import useAsyncEffect from 'use-async-effect';




var parentArray = [];

var sorted = [];
var val = [];
var folders = [];
var users = [];
var groups = [];
var usersGroups = [];
var permission_items = [];
var users_Permission = [];
var roleDefID = [];

// var remainingArr: any = [];
var myVar;
var x;


import 'bootstrap/dist/css/bootstrap.css';
import 'bootstrap/dist/css/bootstrap.min.css';
import { ITreeViewState } from '@pnp/spfx-controls-react/lib/controls/treeView/ITreeViewState';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import { max } from 'lodash';


// import Form from 'react-bootstrap/Form';
// import Button from 'react-bootstrap/Button';

require('./../../../common/css/common.css');
require('./../../../common/css/sidebar.css');
require('./../../../common/css/pagecontent.css');
require('./../../../common/css/spinner.css');


var department;


export default class MyGedTreeView extends React.Component<IMyGedTreeViewProps, IMyGedTreeViewState, any> {

  private graphClient: MSGraphClient;
  readonly context: WebPartContext;

  constructor(props: IMyGedTreeViewProps, context: WebPartContext) {

    super(props, context);


    sp.setup({
      spfxContext: this.props.context
    });

    this.state = {
      selectedKey: null,
      TreeLinks: [],
      parentIDArray: [],
      isLoaded: false
    };

    this.onSelect = this.onSelect.bind(this);



  }

  onSelect(item) {
    this.setState({ selectedKey: item.key });
  }

  private async _getLinks2(sp) {

    var remainingArr: any = [];
    var treearr: any[] = [];
    var testArray: any[] = [];


    var value1 = "TRUE";



    var counterSUB = 0;

    const allItemsMain: any[] = await sp.web.lists.getByTitle('Documents').items
      .select("ID,ParentID,FolderID,Title,IsFolder,description")
      .top(5000)
      .filter("IsFolder eq '" + value1 + "'")
      .getAll();

    testArray = allItemsMain;

    console.log("LENGTH", testArray.length);


    await Promise.all(allItemsMain.map(async (v) => {

      console.log("LOG TEST", v["IsFolder"]);

      if (v["ParentID"] == -1) {

        var str = v["Title"];

        const tree: ITreeItem = {
          id: v["ID"],
          key: v["FolderID"],
          label: str,
          data: 0,
          icon: faFolder,
          children: [],
          revision: "",
          file: "No",
          description: v["description"],
          parentID: v["ParentID"]
        };

        treearr.push(tree);

      }


      else {


        var str = v["Title"];

        const tree: any = {
          id: v["ID"],
          key: v["FolderID"],
          label: str,
          data: 1,
          icon: faFolderOpen,
          revision: "",
          file: "No",
          description: v["description"],
          parentID: v["ParentID"],
          children: []
        };





        // bon la

        var treecol: Array<any> = treearr.filter((value) => {
          return value.key === tree.parentID;
        }).sort((a, b) => {
          if (a.label < b.label) {
            return -1;
          }
          if (a.label > b.label) {
            return 1;
          }
          return 0;
        });


        if (treecol.length != 0) {

          counterSUB = counterSUB + 1;
          treecol[0].children.push(tree);
          treearr.push(tree);
        }

        treearr.push(tree);
      }


    }));



    const sortedTreeArr = treearr.map((tree) => {
      if (tree.children) {
        tree.children.sort((a, b) => a.label.substr(0, 3).localeCompare(b.label.substr(0, 3)));
      }
      return tree;
    }).sort((a, b) => a.label.localeCompare(b.label));


    remainingArr = sortedTreeArr.filter(data => data.key == 1);


    console.log("FOLDERS", allItemsMain.length);

    return remainingArr;

  }

  private getItemId() {
    var queryParms = new URLSearchParams(document.location.search.substring(1));
    var myParm = queryParms.get("folder");
    if (myParm) {
      return myParm.trim();
    }
  }

  async componentDidMount() {
    const x = this.getItemId();

    if (x == null || x == undefined || x == "") {
      const allItems = await this._getLinks2(sp);
      this.setState({ TreeLinks: allItems });
      console.log("COUNT MAIN", allItems);
    } else {
      const parentIDs = await this.getParentID(x);
      const allItems = await this._getLinks2(sp);
      this.setState({ parentIDArray: parentIDs, TreeLinks: allItems });
    }

    // Render the component after _getLinks2() has fully finished
    this.setState({ isLoaded: true });
  }

  private async getParentID(id: any) {

    var parentID = null;
    var value1 = "FALSE";
    var value2 = "TRUE";

    var parentIDArray = [];

    await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + id + "' and IsFolder eq '" + value2 + "'").get().then((results) => {
      parentID = results[0].ParentID;

      // this.setState({ parentIDArray: [...this.state.parentIDArray, parentID] });
      parentIDArray.push(parentID);

      console.log("Parent 1", parentID);

    });


    while (parentID != 1) {
      await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID, Title").filter("FolderID eq '" + parentID + "' and IsFolder eq '" + value2 + "'").get().then((results) => {
        parentID = results[0].ParentID;
        // this.setState({ parentIDArray: [parentID, ...this.state.parentIDArray] });
        parentIDArray.unshift(parentID);

        console.log("Parent 2", parentID);
      });
    }


    // this.setState({ parentIDArray: [...this.state.parentIDArray, parseInt(this.getItemId())] });
    parentIDArray.push(parseInt(this.getItemId()));


    // if (this.state.parentIDArray.length > 1) {

    if (parentIDArray.length > 1) {
      // const temp = this.state.parentIDArray;

      // console.log("TEMP", temp);

      // temp.shift();
      // this.setState({ parentIDArray: [...temp] });

      parentIDArray.shift();
    }

    // parentIDArray.sort(function (a, b) { return a - b });
    console.log("ArrayParent", parentIDArray);

    return parentIDArray;

  }

  public render(): React.ReactElement<IMyGedTreeViewProps> {

    const { TreeLinks, parentIDArray, selectedKey, isLoaded } = this.state;


    var y = [];

    x = this.getItemId();

    // this.getParentID(x);


    console.log("TEST PARENT ARRAY", y);

    console.log("ITEM TO EXPAND", this.getItemId());


    console.log("BEFORE RENDER", this.state.parentIDArray);

    if (!isLoaded) {
      // You can add a loading spinner or a message to show that the component is still loading
      return (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            width: "100%",
            height: "100%",
            backgroundColor: "rgba(0, 0, 0, 0.5)",
            backdropFilter: "blur(5px)",
            zIndex: 9999,
            display: "flex",
            justifyContent: "center",
            alignItems: "center",
          }}
        >
          <img src="https://ncaircalin.sharepoint.com/sites/TestMyGed/SiteAssets/images/loader.gif" alt="Loading..." />
        </div>
      );
    }


    return (

      <div className="container-fluid" style={{ height: "100vh" }}>

        <div className="row" style={{ height: "100vh" }}>
          <div className="col-sm-3">
            <div id="sidebarMenu" className="sidebar">
              <div className="position-sticky">
                <div className="list-group list-group-flush mx-3 mt-4" id="tree">
                  <TreeView
                    items={this.state.TreeLinks}
                    defaultExpanded={true}
                    showCheckboxes={false}
                    treeItemActionsDisplayMode={TreeItemActionsDisplayMode.Buttons}
                    // onSelect={this.onSelect}
                    onExpandCollapse={this.onTreeItemExpandCollapse}
                    onRenderItem={this.renderCustomTreeItem}
                    defaultSelectedKeys={[parseInt(x)]}
                    // defaultSelectedKey={this.state.selectedKey}
                    defaultExpandedKeys={this.state.parentIDArray}
                    expandToSelected={true}
                    defaultExpandedChildren={false}
                    className="my-treeview"
                  />

                </div>
              </div>
            </div>

          </div>

          <div className="col-sm-9">

            <div id="loader"></div>

            <form id="form_metadata">

              <div id="access_form">

                <div className="dossier_headers">
                  <div className="container">
                    <div className="image">
                      <img src='https://ncaircalin.sharepoint.com/sites/TestMyGed/SiteAssets/images/flower.png' />
                      <h2 id='h2_folderName'>
                        Gestion Documentaire
                      </h2>
                    </div>

                  </div>


                  <nav aria-label="breadcrumb" id='nav'>
                    <ul className="breadcrumb" id="folder_nav">
                      <li className="breadcrumb-item"><a href="#" title="Mettre à jour le dossier" role="button" onClick={async (event: React.MouseEvent<HTMLElement>) => { await this.load_folders(); $("#access_rights_form").css("display", "none"); $("#alert_0_doc").css("display", "none"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "block"); $("#doc_details_add").css("display", "none"); }}><FontAwesomeIcon icon={faEdit} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                      <li className="breadcrumb-item"><a href="#" title="Créer un document" role="button" id='add_document' onClick={(event: React.MouseEvent<HTMLElement>) => { $("#access_rights_form").css("display", "none"); $("#alert_0_doc").css("display", "none"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "none"); $("#doc_details_add").css("display", "block"); }}><FontAwesomeIcon icon={faFile} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                      <li className="breadcrumb-item"><a href="#" title="Autorisation sur le dossier" role="button" onClick={async (event: React.MouseEvent<HTMLElement>) => { await this.getSiteUsers(); this.getSiteGroups(); $("#table_documents").css("display", "none"); $("#access_rights_form").css("display", "block"); $("#alert_0_doc").css("display", "none"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "none"); $("#doc_details_add").css("display", "none"); }}><FontAwesomeIcon icon={faLock} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                      <li className="breadcrumb-item"><a href="#" title="Ajouter des sous-dossiers" role="button" onClick={(event: React.MouseEvent<HTMLElement>) => { $("#access_rights_form").css("display", "none"); $("#subfolders_form").css("display", "block"); $("#edit_details").css("display", "none"); $("#alert_0_doc").css("display", "none"); $("#doc_details_add").css("display", "none"); }}><FontAwesomeIcon icon={faFolderPlus} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                      <li className="breadcrumb-item" id='bouton_delete'></li>

                      {/* <li className="breadcrumb-item" id='bouton_delete'><a href="#" title="Supprimer" role="button" id='delete_folder'><FontAwesomeIcon icon={faTrashCan} className="fa-icon fa-2x"></FontAwesomeIcon></a></li> */}
                      <li className="breadcrumb-item"><a href="#" title="Notifier" role="button" onClick={async (event: React.MouseEvent<HTMLElement>) => { await this.getSiteUsers(); this.getSiteGroups(); $("#table_documents").css("display", "none"); $("#access_rights_form").css("display", "none"); $("#alert_0_doc").css("display", "none"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "none"); $("#doc_details_add").css("display", "none"); }} ><FontAwesomeIcon icon={faBell} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                    </ul>
                  </nav>

                </div>




                <h4 id='alert_0_doc'>Veuillez sélectionner un sous répertoire</h4>


                <div id="edit_details">
                  <div className="row">
                    <div className="col-6">
                      <Label>Folder Name
                        <input type="text" className="form-control" id="folder_name1" />
                      </Label>
                    </div>

                    <div className="col-6">
                      <Label>Folder Description
                        <input type="text" className="form-control" id="folder_desc" />
                      </Label>
                    </div>
                  </div>

                  <div className="row">
                    <div className="col-6">
                      <Label>Parent Folder
                        <input type="text" className="form-control" id="parent_folder" list='folders' />

                        <datalist id="folders">
                          <select id="select_folders"></select>
                        </datalist>
                        {/* <select className='form-select' name="parentFolder" id="parent_folder">

            </select> */}
                      </Label>
                    </div>

                    {/* <div className="col-6">
                      <Label>Folder Order
                        <input type="text" className="form-control" id="folder_order" />
                      </Label>
                    </div> */}


                  </div>

                  <div className="row">
                    <div className="col-8">

                    </div>
                    <div className="col-2" id="update_btn_dossier">

                      {/* <button type="button" className="btn btn-primary" id='update_details'>Edit Details</button> */}

                    </div>

                    <div className="col-2">
                      <button type="button" className="btn btn-primary" id='edit_cancel'>Cancel</button>

                    </div>

                  </div>


                </div>

                <div id="access_rights_form">


                  <div className="row">

                    <div className="col-6">
                      <Label>Ajouter un droit d'accès utilisateur

                        <input type="text" className="form-control" id="users_name" list='users' />

                        <datalist id="users">
                          <select id="select_users"></select>
                        </datalist>
                      </Label>
                    </div>


                    <div className="col-3">
                      <Label> Type
                        <select className='form-select' name="permissions" id="permissions_user">
                          <option value="NONE">NONE</option>
                          <option value="READ">READ</option>
                          <option value="READ_WRITE">READ_WRITE</option>
                          <option value="ALL">ALL</option>
                        </select>
                      </Label>
                    </div>
                    <div className="col-3" id="add_btn_user">
                    </div>
                  </div>

                  <div className="row">


                    <div className="col-6">
                      <Label>Ajouter un droit d'accès de groupe
                        <input type="text" className="form-control" id="group_name" list='groups' />

                        <datalist id="groups">
                          <select id="select_groups"></select>
                        </datalist>
                      </Label>
                    </div>


                    <div className="col-3">
                      <Label> Type
                        <select className='form-select' name="permissions" id="permissions_group">
                          <option value="NONE">NONE</option>
                          <option value="READ">READ</option>
                          <option value="READ_WRITE">READ_WRITE</option>
                          <option value="ALL">ALL</option>
                        </select>
                      </Label>
                    </div>
                    <div className="col-3" id="add_btn_group">
                    </div>
                  </div>

                  <div className='row'>
                    <div id="spListPermissions">
                      <table id='tbl_permission' className='table table-striped'>
                        <thead>
                          <tr>
                            <th className="text-left">Nom</th>
                            <th className="text-left" >Droits d'accès</th>
                            <th className="text-left" >Actions</th>
                          </tr>
                        </thead>
                        <tbody id="tbl_permission_bdy">



                        </tbody>
                      </table>
                    </div>
                  </div>

                </div>


                <div id='table_documents'>

                  <div id="spListDocuments">

                    <table id='tbl_documents' className='table table-striped'>
                      <thead>
                        <tr>
                          <th className="text-left" id='nom_doc'>Nom du document</th>
                          <th className="text-left" id='desc_doc'>Description</th>
                          <th className="text-left">Revision</th>
                          <th className="text-right" >Actions</th>
                        </tr>
                      </thead>
                      <tbody id="tbl_documents_bdy">



                      </tbody>
                    </table>
                  </div>


                </div>

                <div id="subfolders_form">
                  <div className="row">
                    <div className="col-6">
                      <Label>Folder name
                        <input type="text" className="form-control" id="folder_name" />
                      </Label>
                    </div>

                    <div className="col-3" id="add_btn_subFolder">

                    </div>
                    <div className="col-3">
                      <button type="button" className="btn btn-primary add_subfolder mb-2 " id="cancel_add_sub">Annuler</button>
                    </div>
                  </div>

                </div>

                {
                  <div id="doc_details_add">

                    <div className="row">

                      <div className="col-6">
                        <Label>Nom du document
                          <input type="text" id='input_doc_number_add' className='form-control' required />
                        </Label>
                      </div>


                      <div className="col-6">
                        <Label>Fichier
                          <input type="file" name="file" id="file_ammendment" className="form-control" />
                        </Label>


                      </div>

                    </div>

                    <div className="row">
                      <div className="col-3">
                        <Label>
                          Revision
                          <input type="text" id='input_revision_add' className='form-control' />
                        </Label>
                      </div>
                      <div className="col-3">
                        <Label>
                          Status
                          <input type="text" id='input_status_add' className='form-control' />
                        </Label>
                      </div>
                      <div className="col-3">
                        <Label>
                          Owner
                          <input type="text" id='input_owner_add' className='form-control' />
                        </Label>
                      </div>
                      <div className="col-3">
                        <Label>
                          Active Date
                          <input type="text" id='input_activeDate_add' className='form-control' />
                        </Label>
                      </div>
                    </div>

                    <div className="row">
                      <div className="col-6">
                        <Label>
                          Filename
                          <input type="text" id='input_filename_add' className='form-control' disabled />
                        </Label>
                      </div>
                      <div className="col-6">
                        <Label>
                          Author
                          <input type="text" id='input_author_add' className='form-control' />
                        </Label>
                      </div>

                    </div>

                    <div className="row">
                      <div className="col-5">
                        <Label>
                          Description
                          <textarea id='input_description_add' className='form-control' rows={2} />
                        </Label>
                      </div>
                      <div className="col-4">
                        <Label>
                          Mots-clés
                          <textarea id='input_keywords_add' className='form-control' rows={2} />
                        </Label>
                      </div>
                      <div className="col-3">
                        <Label>
                          Review Date
                          <input type="text" id='input_reviewDate_add' className='form-control' />
                        </Label>
                      </div>
                    </div>

                    <div className="row">
                      <div className="col-8">

                      </div>
                      <div className="col-2" id="add_document_btn">



                      </div>

                      <div className="col-2">
                        <button type="button" className="btn btn-primary" id='cancel_doc'>Annuler</button>
                      </div>

                    </div>

                  </div>

                }





              </div>

            </form>

          </div>

        </div>
      </div>

    );

  }

  private async load_folders() {



    var value1 = "TRUE";

    var drp_folders = document.getElementById("select_folders");

    // const allItems: any = await sp.web.lists.getByTitle('Documents').items.getAll(),

    const all_folders: any = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID,Title,IsFolder,description").top(5000).filter("IsFolder eq '" + value1 + "'").get();


    // console.log("ALL FOLDERS", all_folders.length);

    folders = all_folders;

    folders.forEach((result: any) => {
      // if(result.IsFolder == "TRUE"){
      // console.log("SELECT_FOLDERS", result.Title);
      var opt = document.createElement('option');
      opt.appendChild(document.createTextNode(result.Title));
      opt.value = result.FolderID + "_" + result.Title;
      drp_folders.appendChild(opt);
      // }

    });

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

  private async addSubfolders(item: ITreeItem) {

    console.log("ID", item.id);
  }

  private async onTreeItemSelect(items: ITreeItem[]) {

    items.forEach((item: any) => {
      $("#h2_folderName").text(item.label);
    });

  }

  // private onSelect(items: ITreeItem[]) {



  //   items.forEach(async (item) => {

  //     $("#h2_folderName").text(item.label);


  //   });



  // }

  private loaderUtil() {
    myVar = setTimeout(this.showForm, 3000);
  }

  private showForm() {
    document.getElementById("loader").style.display = "none";
    document.getElementById("form_metadata").style.display = "block";
  }

  private async onTreeItemExpandCollapse(item: ITreeItem, isExpanded: boolean) {
    console.log((isExpanded ? "Item expanded: " : "Item collapsed: ") + item.label);


    // (isExpanded? + item.)


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

  // private async loadDocs() {

  //   var x = this.getItemId();
  //   var response_doc = null;
  //   var response_distinc = [];
  //   var html_document: string = ``;
  //   var value1 = "FALSE";
  //   var value2 = "TRUE";
  //   var pdfName = '';

  //   const groupTitle = [];
  //   let groups: any = await sp.web.currentUser.groups();

  //   usersGroups = groups;

  //   console.log("USERS GROUPS", usersGroups);

  //   usersGroups.forEach((item) => {

  //     groupTitle.push(item.Title);
  //   });


  //   console.log("DANS NUVO GROUP ARRAY", groupTitle);


  //   // if (groupTitle.includes("myGed Visitors")) {
  //   if (groupTitle.includes("Utilisateur MyGed")) {

  //     $("#nav").css("display", "none");
  //   }
  //   else {

  //     $("#nav").css("display", "block");
  //   }

  //   var document_container: Element = document.getElementById("tbl_documents_bdy");
  //   document_container.innerHTML = "";

  //   const all_documents: any[] = await sp.web.lists.getByTitle('Documents').items
  //     .select("ID,ParentID,FolderID,Title,revision,IsFolder,description, attachmentUrl")
  //     .top(5000)
  //     .filter("ParentID eq '" + parseInt(x) + "' and IsFolder eq '" + value1 + "'")
  //     .getAll();

  //   const dossier: any[] = await sp.web.lists.getByTitle('Documents').items
  //     .select("ID,ParentID,FolderID,Title,revision,IsFolder,description")
  //     .top(5000)
  //     .filter("FolderID eq '" + parseInt(x) + "' and IsFolder eq '" + value2 + "'")
  //     .getAll();

  //   await Promise.all(dossier.map(async (v) => {

  //     $("#h2_folderName").text(v["Title"]);
  //   }));


  //   response_doc = all_documents;

  //   var result = response_doc.filter((obj, pos, arr) => {
  //     return arr.map(mapObj =>
  //       mapObj.Title).lastIndexOf(obj.Title) == pos;
  //   });



  //   //display
  //   {
  //     $("#access_form").css("display", "block");
  //     $("#doc_form").css("display", "none");
  //     $(".dossier_headers").css("display", "block");

  //     $("#subfolders_form").css("display", "none");

  //     $("#access_rights_form").css("display", "none");
  //     $("#notifications_doc_form").css("display", "none");

  //     $("#doc_details_add").css("display", "none");
  //     $("#edit_details").css("display", "none");
  //     // $("#table_documents").css("display", "block");

  //   }


  //   if (result.length > 0) {


  //     html_document = ``;
  //     $("#alert_0_doc").css("display", "none");
  //     $("#table_documents").css("display", "block");

  //     // $("#table_documents").css("display", "block");


  //     await result.forEach(async (element) => {



  //       var externalUrl = '';
  //       var urlFile = '';
  //       html_document += `
  //               <tr>
  //               <td class="text-left">${element.Id}</td>

  //               <td class="text-left">${element.Title}</td>

  //               <td class="text-left"> 
  //               ${element.description}          
  //               </td>


  //               <td>
  //               <a href="#" title="Mettre à jour le document" role="button" id="${element.Id}_view_doc_details" class="btn_view_doc_details">
  //               <svg aria-hidden="true" focusable="false" data-prefix="far" 
  //               data-icon="pen-to-square" class="svg-inline--fa fa-pen-to-square fa-icon fa-2x" 
  //               role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">
  //               <path fill="currentColor" d="M373.1 24.97C401.2-3.147 446.8-3.147 474.9 24.97L487 37.09C515.1 65.21 515.1 110.8 487 138.9L289.8 336.2C281.1 344.8 270.4 351.1 258.6 354.5L158.6 383.1C150.2 385.5 141.2 383.1 135 376.1C128.9 370.8 126.5 361.8 128.9 353.4L157.5 253.4C160.9 241.6 167.2 230.9 175.8 222.2L373.1 24.97zM440.1 58.91C431.6 49.54 416.4 49.54 407 58.91L377.9 88L424 134.1L453.1 104.1C462.5 95.6 462.5 80.4 453.1 71.03L440.1 58.91zM203.7 266.6L186.9 325.1L245.4 308.3C249.4 307.2 252.9 305.1 255.8 302.2L390.1 168L344 121.9L209.8 256.2C206.9 259.1 204.8 262.6 203.7 266.6zM200 64C213.3 64 224 74.75 224 88C224 101.3 213.3 112 200 112H88C65.91 112 48 129.9 48 152V424C48 446.1 65.91 464 88 464H360C382.1 464 400 446.1 400 424V312C400 298.7 410.7 288 424 288C437.3 288 448 298.7 448 312V424C448 472.6 408.6 512 360 512H88C39.4 512 0 472.6 0 424V152C0 103.4 39.4 64 88 64H200z"></path></svg></a>



  //              <a href="#"  title="Voir le document" id="${element.Id}_view_doc" role="button" class="btn_view_doc" style="padding-left: inherit;">
  //              <svg aria-hidden="true" focusable="false" data-prefix="far" 
  //              data-icon="eye" class="svg-inline--fa fa-eye fa-icon fa-2x" 
  //              role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512">
  //              <path fill="currentColor" d="M160 256C160 185.3 217.3 128 288 128C358.7 128 416 185.3 416 256C416 326.7 358.7 384 288 384C217.3 384 160 326.7 160 256zM288 336C332.2 336 368 300.2 368 256C368 211.8 332.2 176 288 176C287.3 176 286.7 176 285.1 176C287.3 181.1 288 186.5 288 192C288 227.3 259.3 256 224 256C218.5 256 213.1 255.3 208 253.1C208 254.7 208 255.3 208 255.1C208 300.2 243.8 336 288 336L288 336zM95.42 112.6C142.5 68.84 207.2 32 288 32C368.8 32 433.5 68.84 480.6 112.6C527.4 156 558.7 207.1 573.5 243.7C576.8 251.6 576.8 260.4 573.5 268.3C558.7 304 527.4 355.1 480.6 399.4C433.5 443.2 368.8 480 288 480C207.2 480 142.5 443.2 95.42 399.4C48.62 355.1 17.34 304 2.461 268.3C-.8205 260.4-.8205 251.6 2.461 243.7C17.34 207.1 48.62 156 95.42 112.6V112.6zM288 80C222.8 80 169.2 109.6 128.1 147.7C89.6 183.5 63.02 225.1 49.44 256C63.02 286 89.6 328.5 128.1 364.3C169.2 402.4 222.8 432 288 432C353.2 432 406.8 402.4 447.9 364.3C486.4 328.5 512.1 286 526.6 256C512.1 225.1 486.4 183.5 447.9 147.7C406.8 109.6 353.2 80 288 80V80z">
  //              </path></svg>

  //              </a>

  //               </td>

  //              `;

  //              externalUrl = element.attachmentUrl;


  //       await sp.web.lists.getByTitle("Documents")
  //         .items
  //         .getById(parseInt(element.Id))
  //         .attachmentFiles
  //         .select('FileName', 'ServerRelativeUrl')
  //         .get()
  //         .then(responseAttachments => {
  //           responseAttachments
  //             .forEach(attachmentItem => {
  //               pdfName = attachmentItem.FileName;
  //               urlFile = attachmentItem.ServerRelativeUrl;
  //             });

  //         })


  //         .then(async () => {



  //           const btn_view_doc = document.getElementById(element.Id + '_view_doc');
  //           const btn_view_doc_details = document.getElementById(element.Id + '_view_doc_details');


  //           await btn_view_doc?.addEventListener('click', async (event) => {


  //             $(".modal").css("display", "block");

  //             if(externalUrl !== ''){
  //               alert("EXTERNAL FILE IS " +externalUrl);
  //               window.open(`${externalUrl}`, '_blank');

  //             }
  //             else{
  //               alert("EXTERNAL FILE IS NULL");
  //               window.open(`${urlFile}`, '_blank');
  //             }

  //           });

  //           await btn_view_doc_details?.addEventListener('click', async () => {

  //             window.open(`https://ncaircalin.sharepoint.com/sites/TestMyGed/SitePages/Document.aspx?document=${element.Title}&documentId=${element.FolderID}`, '_blank');

  //           });

  //         });

  //       console.log("URL FILE", urlFile);


  //     });

  //     document_container.innerHTML += html_document;

  //   }

  //   else {
  //     $("#alert_0_doc").css("display", "block");
  //     $("#table_documents").css("display", "none");
  //   }
  // }

  private renderCustomTreeItem(item: ITreeItem): JSX.Element {

    function solid(arg0: string): IconProp {
      throw new Error('Function not implemented.');
    }

    return (
      <span

        onClick={async (event: React.MouseEvent<HTMLInputElement>) => {
          const groupTitle = [];
          let groups: any = await sp.web.currentUser.groups();

          usersGroups = groups;

          console.log("USERS GROUPS", usersGroups);

          usersGroups.forEach((item) => {

            groupTitle.push(item.Title);
          });


          console.log("DANS NUVO GROUP ARRAY", groupTitle);


          if (groupTitle.includes("Utilisateur MyGed")) {

            $("#nav").css("display", "none");
          }
          else {

            $("#nav").css("display", "block");
          }

          console.log("GROOOOUP", groups);

          //display
          {
            $("#access_form").css("display", "block");
            $("#doc_form").css("display", "none");
            $(".dossier_headers").css("display", "block");

            $("#subfolders_form").css("display", "none");

            $("#access_rights_form").css("display", "none");
            $("#notifications_doc_form").css("display", "none");

            $("#doc_details_add").css("display", "none");
            $("#edit_details").css("display", "none");
            $("#h2_folderName").text(item.label);
          }

          $("#h2_folderName").text(item.label);

          //render table
          {

            var response_doc = null;
            var response_distinc = [];
            var html_document: string = ``;
            var value1 = "FALSE";
            var value2 = "TRUE";
            var value3 = "";


            var pdfName = '';

            console.log("ITEM KEY", item.key);


            var document_container: Element = document.getElementById("tbl_documents_bdy");

            document_container.innerHTML = '';


            const all_documents: any[] = await sp.web.lists.getByTitle('Documents').items
              .select("ID,ParentID,FolderID,Title,revision,IsFolder,description, attachmentUrl")
              .top(5000)
              .filter("ParentID eq '" + parseInt(item.key) + "' and IsFolder eq '" + value1 + "' and revision ne '" + value3 + "' ")
              .getAll();


            console.log("CLICK LENGTH", all_documents.length);
            console.log("CLICK LENGTH", all_documents);


            response_doc = all_documents;

            var result = response_doc.filter((obj, pos, arr) => {
              return arr.map(mapObj =>
                mapObj.Title).lastIndexOf(obj.Title) == pos;
            }).sort((a, b) => (a.Title > b.Title) ? 1 : -1);



            console.log("ALL", response_doc);


            console.log("RESULT DISTINCT", result);
            console.log("RESULT DISTINCT ARRAY LOT LA", response_distinc);


            if (result.length > 0) {


              html_document = ``;
              $("#alert_0_doc").css("display", "none");
              $("#table_documents").css("display", "block");



              await result.forEach(async (element) => {


                if (element.revision !== null || element.revision !== undefined || element.revision !== "") {

                  var urlFile = '';
                  var externalFileUrl = element.attachmentUrl;
                  html_document += `
                  <tr>
  
                  <td class="text-left">${element.Title}</td>
  
                  <td class="text-left"> 
                  ${element.description}          
                  </td>

                  <td class="text-left">${element.revision}</td>
  
                  
                  <td style="font-size: 8px;">
                  <a href="#" title="Mettre à jour le document" role="button" id="${element.Id}_view_doc_details" class="btn_view_doc_details" style="text-decoration: auto;padding-right: 1em;">
                  <svg aria-hidden="true" focusable="false" data-prefix="far" 
                 data-icon="eye" class="svg-inline--fa fa-eye fa-icon fa-2x" 
                 role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 288 256"><!--! Font Awesome Pro 6.3.0 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license (Commercial License) Copyright 2023 Fonticons, Inc. --><path d="M256 512A256 256 0 1 0 256 0a256 256 0 1 0 0 512zM216 336h24V272H216c-13.3 0-24-10.7-24-24s10.7-24 24-24h48c13.3 0 24 10.7 24 24v88h8c13.3 0 24 10.7 24 24s-10.7 24-24 24H216c-13.3 0-24-10.7-24-24s10.7-24 24-24zm40-208a32 32 0 1 1 0 64 32 32 0 1 1 0-64z"/></svg>
                  </a>
  
  
  
                 <a href="#"  title="Voir le document" id="${element.Id}_view_doc" role="button"  class="btn_view_doc" style="padding-left: inherit;">
                 <svg aria-hidden="true" focusable="false" data-prefix="far" 
                 data-icon="eye" class="svg-inline--fa fa-eye fa-icon fa-2x" 
                 role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 288 256">
                 <path fill="currentColor" d="M160 256C160 185.3 217.3 128 288 128C358.7 128 416 185.3 416 256C416 326.7 358.7 384 288 384C217.3 384 160 326.7 160 256zM288 336C332.2 336 368 300.2 368 256C368 211.8 332.2 176 288 176C287.3 176 286.7 176 285.1 176C287.3 181.1 288 186.5 288 192C288 227.3 259.3 256 224 256C218.5 256 213.1 255.3 208 253.1C208 254.7 208 255.3 208 255.1C208 300.2 243.8 336 288 336L288 336zM95.42 112.6C142.5 68.84 207.2 32 288 32C368.8 32 433.5 68.84 480.6 112.6C527.4 156 558.7 207.1 573.5 243.7C576.8 251.6 576.8 260.4 573.5 268.3C558.7 304 527.4 355.1 480.6 399.4C433.5 443.2 368.8 480 288 480C207.2 480 142.5 443.2 95.42 399.4C48.62 355.1 17.34 304 2.461 268.3C-.8205 260.4-.8205 251.6 2.461 243.7C17.34 207.1 48.62 156 95.42 112.6V112.6zM288 80C222.8 80 169.2 109.6 128.1 147.7C89.6 183.5 63.02 225.1 49.44 256C63.02 286 89.6 328.5 128.1 364.3C169.2 402.4 222.8 432 288 432C353.2 432 406.8 402.4 447.9 364.3C486.4 328.5 512.1 286 526.6 256C512.1 225.1 486.4 183.5 447.9 147.7C406.8 109.6 353.2 80 288 80V80z">
                 </path></svg>
  
                 </a>
  
                  </td>
                
                 `;


                  await sp.web.lists.getByTitle("Documents")
                    .items
                    .getById(parseInt(element.Id))
                    .attachmentFiles
                    .select('FileName', 'ServerRelativeUrl')
                    .get()
                    .then(responseAttachments => {
                      responseAttachments
                        .forEach(attachmentItem => {
                          pdfName = attachmentItem.FileName;
                          urlFile = attachmentItem.ServerRelativeUrl;
                        });

                    })


                    .then(async () => {

                      const btn_view_doc = document.getElementById(element.Id + '_view_doc');
                      const btn_view_doc_details = document.getElementById(element.Id + '_view_doc_details');

                      await btn_view_doc?.addEventListener('click', async (event) => {

                        $(".modal").css("display", "block");

                        if (externalFileUrl == undefined || externalFileUrl == null || externalFileUrl == "") {
                          window.open(`${urlFile}`, '_blank');
                        }
                        else {
                          window.open(`${externalFileUrl}`, '_blank');
                        }

                      });

                      //view details_doc
                      await btn_view_doc_details?.addEventListener('click', async () => {
                        window.open(`https://ncaircalin.sharepoint.com/sites/TestMyGed/SitePages/Document.aspx?document=${element.Title}&documentId=${element.FolderID}`, '_blank');
                      });

                      $("#edit_cancel").click(() => {

                        $("#edit_details").css("display", "none");

                      });

                    });

                  console.log("URL FILE", urlFile);

                }

              });

              document_container.innerHTML += html_document;

            }

            else {
              $("#alert_0_doc").css("display", "block");
              $("#table_documents").css("display", "none");
            }

            // $("#tbl_documents").DataTable({
            //   columnDefs: [
            //     {
            //       target: 0, // targets the second and fourth columns
            //       width: '15%' // sets the width of the columns to 20% of the table's width
            //     }
            //     ,
            //     {
            //       target: 1, // targets the second and fourth columns
            //       width: '60%' // sets the width of the columns to 20% of the table's width
            //     }
            //     ,
            //     {
            //       target: 2, // targets the second and fourth columns
            //       width: '8%' // sets the width of the columns to 20% of the table's width
            //     }
            //   ]
            // });
          }

          //render metadata
          {
            var fileName = "";
            var content = null;

            var filename_add = "";
            var content_add = null;

            var titleFolder = "";

            const allItemsFolder: any[] = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID,Title,revision,IsFolder,description").filter("FolderID eq '" + item.parentID + "'").getAll();

            allItemsFolder.forEach((x) => {

              titleFolder = x.Title;

            });

            $("#folder_name1").val(item.label);
            $("#folder_desc").val(item.description);
            $("#parent_folder").val(item.parentID + "_" + titleFolder);
          }

          //bouton delete dossier
          {
            var delete_dossier: Element = document.getElementById("bouton_delete");


            let nav_html_delete_dossier: string = '';


            // console.log("ONSELECT", item.label);

            nav_html_delete_dossier = `
                          <a href="#" title="Supprimer" 
                          role="button" id='${item.id}_deleteFolder'>
                      <svg aria-hidden="true" focusable="false" data-prefix="far" data-icon="trash-can" 
                      class="svg-inline--fa fa-trash-can fa-icon fa-2x" role="img" xmlns="http://www.w3.org/2000/svg" 
                      viewBox="0 0 448 512">
                      <path fill="currentColor" d="M160 400C160 408.8 152.8 416 144 416C135.2 416 128 408.8 128 400V192C128 183.2 135.2 176 144 176C152.8 176 160 183.2 160 192V400zM240 400C240 408.8 232.8 416 224 416C215.2 416 208 408.8 208 400V192C208 183.2 215.2 176 224 176C232.8 176 240 183.2 240 192V400zM320 400C320 408.8 312.8 416 304 416C295.2 416 288 408.8 288 400V192C288 183.2 295.2 176 304 176C312.8 176 320 183.2 320 192V400zM317.5 24.94L354.2 80H424C437.3 80 448 90.75 448 104C448 117.3 437.3 128 424 128H416V432C416 476.2 380.2 512 336 512H112C67.82 512 32 476.2 32 432V128H24C10.75 128 0 117.3 0 104C0 90.75 10.75 80 24 80H93.82L130.5 24.94C140.9 9.357 158.4 0 177.1 0H270.9C289.6 0 307.1 9.358 317.5 24.94H317.5zM151.5 80H296.5L277.5 51.56C276 49.34 273.5 48 270.9 48H177.1C174.5 48 171.1 49.34 170.5 51.56L151.5 80zM80 432C80 449.7 94.33 464 112 464H336C353.7 464 368 449.7 368 432V128H80V432z"></path></svg> 
                          </a>
                          
                          `;

            delete_dossier.innerHTML = nav_html_delete_dossier;

            const btn = document.getElementById(item.id + '_deleteFolder');

            await btn?.addEventListener('click', async () => {
              // this.domElement.querySelector('#btn' + item.Id + '_edit').addEventListener('click', () => {
              //localStorage.setItem("contractId", item.Id);
              if (confirm(`Êtes-vous sûr de vouloir supprimer ${item.label} ?`)) {

                try {
                  var res = await sp.web.lists.getByTitle('Documents').items.getById(parseInt(item.id)).delete()
                    .then(() => {
                      alert("Dossier supprimé avec succès.");
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

            });


            $("#edit_cancel").click(() => {

              $("#edit_details").css("display", "none");
            });

          }

          //bouton update dossier
          {
            var update_dossier_container: Element = document.getElementById("update_btn_dossier");

            let update_btn_dossier: string = `<button type="button" class="btn btn-primary btn_edit_dossier" id='${item.id}_update_details'>Edit Details</button>
          `;

            update_dossier_container.innerHTML = update_btn_dossier;


            const btn_edit_dossier = document.getElementById(item.id + '_update_details');

            await btn_edit_dossier?.addEventListener('click', async () => {


              let text = $("#parent_folder").val();
              const myArray = text.toString().split("_");
              let parentId = myArray[0];

              if (confirm(`Etes-vous sûr de vouloir mettre à jour les détails de ${item.label} ?`)) {

                try {

                  const i = await await sp.web.lists.getByTitle('Documents').items.getById(parseInt(item.id)).update({
                    Title: $("#folder_name1").val(),
                    description: $("#folder_desc").val(),
                    ParentID: parseInt(parentId)

                  })
                    .then(() => {

                      alert("Détails mis à jour avec succès");
                    })
                    .then(() => {

                      window.open(`https://ncaircalin.sharepoint.com/sites/TestMyGed/SitePages/Home.aspx?folder=${item.key}`, "blank");
                    });

                }
                catch (err) {
                  alert(err.message);
                }

              }
              else {

              }

            });
          }

          //bouton upload document
          {
            var add_doc_container: Element = document.getElementById("add_document_btn");

            let add_btn_document: string = `
          <button type="button" class="btn btn-primary add_doc" id=${item.id}_add_doc>Sauvegarder</button>
          `;

            add_doc_container.innerHTML = add_btn_document;


            const btn_add_doc = document.getElementById(item.id + '_add_doc');

            await btn_add_doc?.addEventListener('click', async () => {

              let user_current = await sp.web.currentUser();

              console.log("CURRENT USER", user_current);


              if ($('#file_ammendment').val() == '') {

                alert("Veuillez télécharger le fichier avant de continuer.");

              }
              else {

                if (confirm(`Etes-vous sûr de vouloir creer un document ?`)) {


                  try {

                    const i = await await sp.web.lists.getByTitle('Documents').items.add({
                      Title: $("#input_doc_number_add").val(),
                      description: $("#input_description_add").val(),
                      doc_number: $("#input_doc_number_add").val(),
                      revision: $("#input_revision_add").val(),
                      ParentID: item.key,
                      IsFolder: "FALSE",
                      keywords: $("#input_keywords_add").val(),
                      owner: user_current.Title,
                      createdDate: new Date().toLocaleString()
                    })
                      .then(async (iar) => {

                        item = iar.data.ID;


                        const list = sp.web.lists.getByTitle("Documents");

                        await list.items.getById(iar.data.ID).attachmentFiles.add(fileName, content)
                          .then(async () => {

                            await list.items.getById(iar.data.ID).update({
                              FolderID: parseInt(iar.data.ID),
                              filename: fileName
                            });

                            try {
                              // response_same_doc.forEach(async (x) => {

                              await sp.web.lists.getByTitle("Audit").items.add({
                                Title: iar.data.Title.toString(),
                                DateCreated: moment().format("MM/DD/YYYY HH:mm:ss"),
                                Action: "Creation",
                                FolderID: iar.data.ID.toString(),
                                Person: user_current.Title.toString()
                              });
                            }

                            catch (e) {
                              alert("Erreur: " + e.message);
                            }

                          });

                      })
                      .then(() => {

                        alert("Document creer avec succès");
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


            });

          }

          //bouton add subfolder
          {
            var add_subfolder_container: Element = document.getElementById("add_btn_subFolder");

            let add_btn_subfolder: string = `
          <button type="button" class="btn btn-primary add_subfolder mb-2" id="${item.id}_add_btn_subfolder" style="float: right;">Add subfolder</button>
          `;

            add_subfolder_container.innerHTML = add_btn_subfolder;



            const btn_add_subfolder = document.getElementById(item.id + '_add_btn_subfolder');


            await btn_add_subfolder?.addEventListener('click', async () => {
              var subId = null;

              try {
                await sp.web.lists.getByTitle("Documents").items.add({
                  Title: $("#folder_name").val(),
                  ParentID: item.key,
                  IsFolder: "TRUE"
                })
                  .then(async (iar) => {

                    const list = sp.web.lists.getByTitle("Documents");

                    subId = iar.data.ID;

                    await list.items.getById(iar.data.ID).update({
                      FolderID: parseInt(iar.data.ID),


                    })
                      .then(() => {

                        alert(`Dossier ajouté avec succès`);
                      })
                      .then(() => {

                        // window.open("https://ncaircalin.sharepoint.com/sites/TestMyGed/SitePages/Home.aspx?folder=" + subId)
                        window.location.reload();
                      });

                  });

              }
              catch (err) {
                console.log("Erreur:", err.message);
              }


            });

            $("#cancel_add_sub").click(() => {

              $("#subfolders_form").css("display", "none");

            });




          }

          //upload file for new
          {
            $('#file_ammendment').on('change', () => {
              const input = document.getElementById('file_ammendment') as HTMLInputElement | null;


              var file = input.files[0];
              var reader = new FileReader();

              reader.onload = ((file1) => {
                return (e) => {
                  console.log(file1.name);

                  fileName = file1.name,
                    content = e.target.result

                  $("#input_filename_add").val(file1.name);

                };
              })(file);

              reader.readAsArrayBuffer(file);
            });
          }

          //upload file for update
          {
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

          //azoute permission
          {
            //add permission user



            var add_user_permission_container: Element = document.getElementById("add_btn_user");

            let add_btn_user_permission: string = `
      <button type="button" class="btn btn-primary add_group mb-2" id=${item.id}_add_user>Ajouter</button>
      `;

            add_user_permission_container.innerHTML = add_btn_user_permission;

            const btn_add_user = document.getElementById(item.id + '_add_user');

            var peopleID = null;


            await btn_add_user?.addEventListener('click', async () => {

              const user: any = await sp.web.siteUsers.getByEmail($("#users_name").val().toString())();

              users_Permission = user;

              console.log("USERS FOR PERMISSION", users_Permission);

              try {
                console.log("KEY", item.key);

                await sp.web.lists.getByTitle("AccessRights").items.add({
                  Title: item.label.toString(),
                  groupName: $("#users_name").val(),
                  permission: $("#permissions_user option:selected").val(),
                  FolderIDId: item.id.toString(),
                  PrincipleID: user.Id
                })
                  .then(() => {
                    alert("Autorisation ajoutée à ce dossier avec succès.")
                  })
                  .then(() => {
                    window.location.reload();
                  });
              }

              catch (e) {
                alert("Erreur: " + e.message);
              }


              // }


            });
          }



          //close doc upload
          {
            $("#cancel_doc").click(() => {

              $("#doc_details_add").css("display", "none");
            });
          }

          //permission table 
          //load table permission

          {
            var response = null;
            let html: string = ``;

            var permission_container: Element = document.getElementById("tbl_permission");
            permission_container.innerHTML = "";


            const allPermissions: any[] = await sp.web.lists.getByTitle('AccessRights').items.select("ID,groupName,permission,FolderIDId").filter("FolderIDId eq '" + item.id + "'").getAll();

            response = allPermissions;

            console.log(response);

            if (response.length > 0) {
              await response.forEach(element => {

                html += `
                               <tr>
                               <td class="text-left">${element.groupName}</td>
                               
                               <td class="text-left"> 
                               <input type="text" className="form-control" id="permission_value" list='perm' value='${element.permission}'/>
                               
                               
                               <datalist id="perm">
                               <select class='form-select' name="permissions_render" id="permissions_user_render">
                               
                               
                               <option value="NONE">NONE</option>
                               <option value="READ">READ</option>
                               <option value="READ_WRITE">READ_WRITE</option>
                               <option value="ALL">ALL</option>
                               </select>
                               
                               </datalist>
                               
                               </td>
                               
                               <td>
                               <button id="btn${element.ID}_edit" class='buttoncss' type="button">CHANGER</button>
                               
                               
                               </td>
                               </tr>
                               `;

              });


              permission_container.innerHTML += html;

              $("#spListPermissions").css("display", "block");


            }
            else {



            }


          }

        }

        }
      >

        {
          < FontAwesomeIcon icon={item.icon} className="fa-icon" ></FontAwesomeIcon >
        }

        &nbsp;

        {item.label}

      </span>
    );


  }



}


