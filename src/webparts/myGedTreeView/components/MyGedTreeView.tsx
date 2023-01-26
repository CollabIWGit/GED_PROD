import * as React from 'react';
import styles from './MyGedTreeView.module.scss';
import { MSGraphClient } from '@microsoft/sp-http';
import { IMyGedTreeViewProps, IMyGedTreeViewState } from './IMyGedTreeView';
import { TreeView, ITreeItem, TreeViewSelectionMode, TreeItemActionsDisplayMode } from "@pnp/spfx-controls-react/lib/TreeView";
import 'bootstrap/dist/css/bootstrap.min.css';
import $ from 'jquery';
import Popper from 'popper.js';
import 'bootstrap/dist/js/bootstrap.bundle.min';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { sp, List, IItemAddResult, UserCustomActionScope, Items, Item, ITerm } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { getIconClassName, Label } from 'office-ui-fabric-react';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faFolder, faFolderOpen, faFileWord, faEdit, faTrashCan, faBell, faEye } from '@fortawesome/free-regular-svg-icons'
import { faFile, faLock, faFolderPlus, faDownload } from '@fortawesome/free-solid-svg-icons'
import { icon, IconName, IconProp, parse } from '@fortawesome/fontawesome-svg-core';
import { useEffect, useState } from 'react';
//import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { IAttachmentInfo } from "@pnp/sp/attachments";
import "@pnp/sp/attachments";
import { IItem } from "@pnp/sp/items/types";
// import Form from 'react-bootstrap-form';
import * as sharepointConfig from './../../../common/utils/sharepoint-config.json';






var parentIDArray = [];
var sorted = [];
var val = [];
var folders = [];



import 'bootstrap/dist/css/bootstrap.css';
import 'bootstrap/dist/css/bootstrap.min.css';
import { ITreeViewState } from '@pnp/spfx-controls-react/lib/controls/treeView/ITreeViewState';


// import Form from 'react-bootstrap/Form';
// import Button from 'react-bootstrap/Button';

require('./../../../common/css/common.css');
require('./../../../common/css/sidebar.css');
require('./../../../common/css/pagecontent.css');
require('./../../../common/css/spinner.css');

var department;


export default class MyGedTreeView extends React.Component<IMyGedTreeViewProps, IMyGedTreeViewState, any> {

  private graphClient: MSGraphClient;

  constructor(props: IMyGedTreeViewProps, context) {

    super(props, context);


    sp.setup({
      spfxContext: this.props.context
      //props.context
    });




    var x = this.getItemId();

    this.getParentID(x);

    // const sp = spfi().using(SPFx(this.props.context));
    this.state = {
      TreeLinks: [],
    };


    this._getLinks2(sp);

    // this._getLinks3(sp);  //sa pu tester doc library

    // this._getLinks(sp);
    this.render();

    $("#h2_folderName").text("");






  }


  private async _getLinks2(sp) {

    // var treearr: ITreeItem[] = [];
    var treearr: any[] = [];


    //var treearr;
    var treeSub: ITreeItem[] = [];
    var tree: ITreeItem[] = [];

    var value1 = "TRUE";
    var value2 = "FALSE";

    var keysMissing: any = [];
    var allKeys: any = [];
    var keysPresent: any = [];

    var counter = 0;
    var counterSUB = 0;


    const allItemsMain: any[] = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID,Title,IsFolder,description").top(5000).filter("IsFolder eq '" + value1 + "'").get();
    const allItemsFile: any[] = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID,Title,revision,IsFolder,description").filter("IsFolder eq '" + value2 + "'").getAll();


    //const allItemsMain_sorted: any[] = allItemsMain.sort((a, b) => { return a.Title - b.Title });
    // const allItemsMain_sorted: any[] = allItemsMain.sort();


    console.log("ARRRANGED ARRAY: " + allItemsMain.length);

    // const allItemsMain_sorted: any[] = allItemsMain.sort((a, b) => a.FolderID - b.FolderID);
    var x = 0;



    allItemsMain.forEach(v => {


      if (v["ParentID"] == -1) {

        // console.log("We have the main folder here.");

        var str = v["Title"];

        const tree: ITreeItem = {
          // v.id ==> v.FolderID
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
        // console.log("Tree 1", tree);

      }

      // v["FileSystemObjectType"] ==> v["IsFolder"]
      else {

        allKeys.push(v["FolderID"]);

        // console.log("We have a sub folder here.");
        var str = v["Title"];

        const tree: any = {
          // v.id ==> v.FolderID
          id: v["ID"],
          key: v["FolderID"],
          // key: v["Title"],
          // label: str.normalize('NFD').replace(/\p{Diacritic}/gu, ""),
          label: str,
          data: 1,
          icon: faFolderOpen,
          revision: "",
          file: "No",
          description: v["description"],
          parentID: v["ParentID"],
          children: []
        };

        // console.log("Tree 2", tree);

        // ParentID ==> FolderID

        //  var treecol: Array<any> = treearr.filter((value) => { return value.key === v["ParentID"]; }).sort((a, b) => {return a.label - b.label} );

        // var treecol: Array<any> = treearr.filter((value) => { return value.key === v["ParentID"]; });
        var treecol: Array<any> = treearr.filter((value) => { return value.key === tree.parentID; });

        treecol.forEach(() => {

          // console.log("TREEEEEEEECCCCOOOOL", treecol[0].label);
        });

        //  if (treecol.length != 0) {
        if (treecol.length != 0) {

          counterSUB = counterSUB + 1;
          keysPresent.push(tree.key);
          treecol[0].children.push(tree);
          // console.log("TREE COL", treecol);
          // console.log("COL SUB", treeSub);
          treearr.push(tree);
        }

        treearr.push(tree);
      }

    });


    keysMissing = allKeys
      .filter(x => !keysPresent.includes(x))
      .concat(keysPresent.filter(x => !allKeys.includes(x)));

    keysMissing.forEach(v => {

      allItemsMain.forEach(x => {

        if (v === x["FolderID"]) {

          var str = x["Title"];

          const tree: ITreeItem = {
            // v.id ==> v.FolderID
            id: v["ID"],
            key: v["FolderID"],
            // key: v["Title"],
            // label: str.normalize('NFD').replace(/\p{Diacritic}/gu, ""),
            label: str,
            data: 1,
            icon: faFolderOpen,
            children: [],
            revision: "",
            file: "No",
            description: v["description"],
            parentID: v["ParentID"]

          };

          var treecol: Array<ITreeItem> = treearr.filter((value) => { return value.key === x["ParentID"]; });

          if (treecol.length != 0) {
            keysPresent.push(tree.key);
            treecol[0].children.push(tree);
          }

          treearr.push(tree);

        }
      });

    });


    // console.log("KEYS MISSING LENGTH", keysMissing.length);
    // console.log("KEYS MISSING", keysMissing);
    // console.log("KEYS PRESENT LENGTH", keysPresent.length);
    // console.log("KEYS MISSING LENGTH", 763 - keysPresent.length);
    // console.log("COUNTER", counter);
    // console.log("COUNTERSUB", counterSUB);
    // console.log("TREE ARRAY LENGTH", treearr.length);

    allItemsFile.forEach((v) => {
      // console.log("We have a file here.");


      //v["ServerRedirectedEmbedUrl"] ==> v["FileUrl"]
      const iconName: IconProp = faFile;

      const tree: ITreeItem = {
        // v.id ==> v.FolderID
        id: v["ID"],
        //    key: v["FolderID"],
        key: v["FolderID"],
        label: v["Title"] + "-" + v["revision"],
        icon: faFile,
        data: v["FileUrl"],
        revision: v['revision'],
        file: "Yes",
        description: v["description"],
        parentID: v["ParentID"]

      };


      // ParentID ==> FolderID

      var treecol: Array<ITreeItem> = treearr.filter((value) => { return value.file === "No" && value.key == v["ParentID"]; });

      if (treecol.length != 0) {
        treecol[0].children.push(tree);
      }

    });

    // var remainingArr = treearr.filter(data => data.data === 0);

    var remainingArr = treearr.filter(data => data.key == 1);

    //   var finalArr = remainingArr.sort((a, b) =>{
    //     return a.label - b.label;
    // });


    console.log("REMAINING ARRAY ", remainingArr);
    console.log("REMAINING ARRAY LENGTH", remainingArr.length);

    //  Array.prototype.push.apply(allItemsFile, allItemsMain);
    // const mergedArray = [ ...allItemsFile, ...allItemsMain ];

    const unique = [];

    // mergedArray.map(x => unique.filter(a => a.key == x.key).length > 0 ? null : unique.push(x));

    // all.map(x => unique.filter(a => a.key == x.key).length > 0 ? null : unique.push(x));

    // console.log(unique);

    // tree = remainingArr;

    // this.setState({ TreeLinks: remainingArr });
    this.setState({ TreeLinks: remainingArr });
    console.log("FOLDERS", allItemsMain.length);
    // console.log("FILES", allItemsFile.length);
    // console.log("MERGED", mergedArray.length);



  }

  //testing doc library
  private async _getLinks3(sp) {

    // var treearr: ITreeItem[] = [];
    var treearr: any[] = [];


    //var treearr;
    var treeSub: ITreeItem[] = [];
    var tree: ITreeItem[] = [];

    var value1 = "TRUE";
    var value2 = "FALSE";

    var keysMissing: any = [];
    var allKeys: any = [];
    var keysPresent: any = [];

    var counter = 0;
    var counterSUB = 0;


    const allItemsMain: any[] = await sp.web.lists.getByTitle('Test').items.select("ID,ParentID,FolderID,Title,IsFolder").top(5000).filter("IsFolder eq '" + value1 + "'").get();
    const allItemsFile: any[] = await sp.web.lists.getByTitle('Test').items.select("ID,ParentID,FolderID,Title,IsFolder").filter("IsFolder eq '" + value2 + "'").get();


    //const allItemsMain_sorted: any[] = allItemsMain.sort((a, b) => { return a.Title - b.Title });
    // const allItemsMain_sorted: any[] = allItemsMain.sort();


    console.log("ARRRANGED ARRAY: " + allItemsMain.length);

    // const allItemsMain_sorted: any[] = allItemsMain.sort((a, b) => a.FolderID - b.FolderID);
    var x = 0;



    allItemsMain.forEach(v => {


      if (v["ParentID"] == -1) {

        // console.log("We have the main folder here.");

        var str = v["Title"];

        const tree: ITreeItem = {
          // v.id ==> v.FolderID
          id: v["ID"],
          key: v["FolderID"],
          label: str,
          data: 0,
          icon: faFolder,
          children: [],
          // revision: "",
          // file: "No",
          //description: v["description"],
          parentID: v["ParentID"]
        };

        treearr.push(tree);
        // console.log("Tree 1", tree);

      }

      // v["FileSystemObjectType"] ==> v["IsFolder"]
      else {

        allKeys.push(v["FolderID"]);

        // console.log("We have a sub folder here.");
        var str = v["Title"];

        const tree: any = {
          // v.id ==> v.FolderID
          id: v["ID"],
          key: v["FolderID"],
          // key: v["Title"],
          // label: str.normalize('NFD').replace(/\p{Diacritic}/gu, ""),
          label: str,
          data: 1,
          icon: faFolderOpen,
          // revision: "",
          // file: "No",
          // description: v["description"],
          parentID: v["ParentID"],
          children: []
        };

        // console.log("Tree 2", tree);

        // ParentID ==> FolderID

        //  var treecol: Array<any> = treearr.filter((value) => { return value.key === v["ParentID"]; }).sort((a, b) => {return a.label - b.label} );

        // var treecol: Array<any> = treearr.filter((value) => { return value.key === v["ParentID"]; });
        var treecol: Array<any> = treearr.filter((value) => { return value.key === tree.parentID; });

        treecol.forEach(() => {

          // console.log("TREEEEEEEECCCCOOOOL", treecol[0].label);
        });

        //  if (treecol.length != 0) {
        if (treecol.length != 0) {

          counterSUB = counterSUB + 1;
          keysPresent.push(tree.key);
          treecol[0].children.push(tree);
          // console.log("TREE COL", treecol);
          // console.log("COL SUB", treeSub);
          treearr.push(tree);
        }

        treearr.push(tree);
      }

    });


    keysMissing = allKeys
      .filter(x => !keysPresent.includes(x))
      .concat(keysPresent.filter(x => !allKeys.includes(x)));

    keysMissing.forEach(v => {

      allItemsMain.forEach(x => {

        if (v === x["FolderID"]) {

          var str = x["Title"];

          const tree: ITreeItem = {
            // v.id ==> v.FolderID
            id: v["ID"],
            key: v["FolderID"],
            // key: v["Title"],
            // label: str.normalize('NFD').replace(/\p{Diacritic}/gu, ""),
            label: str,
            data: 1,
            icon: faFolderOpen,
            children: [],
            // revision: "",
            // file: "No",
            // description: v["description"],
            parentID: v["ParentID"]

          };

          var treecol: Array<ITreeItem> = treearr.filter((value) => { return value.key === x["ParentID"]; });

          if (treecol.length != 0) {
            keysPresent.push(tree.key);
            treecol[0].children.push(tree);
          }

          treearr.push(tree);

        }
      });

    });


    // console.log("KEYS MISSING LENGTH", keysMissing.length);
    // console.log("KEYS MISSING", keysMissing);
    // console.log("KEYS PRESENT LENGTH", keysPresent.length);
    // console.log("KEYS MISSING LENGTH", 763 - keysPresent.length);
    // console.log("COUNTER", counter);
    // console.log("COUNTERSUB", counterSUB);
    // console.log("TREE ARRAY LENGTH", treearr.length);

    allItemsFile.forEach((v) => {
      console.log("We have a file here.");


      //v["ServerRedirectedEmbedUrl"] ==> v["FileUrl"]
      const iconName: IconProp = faFile;

      const tree: ITreeItem = {
        // v.id ==> v.FolderID
        id: v["ID"],
        //    key: v["FolderID"],
        key: v["FolderID"],
        label: v["Title"],
        icon: faFile,
        data: v["FileUrl"],
        // revision: v['revision'],
        // file: "Yes",
        // description: v["description"],
        parentID: v["ParentID"]

      };


      // ParentID ==> FolderID

      var treecol: Array<ITreeItem> = treearr.filter((value) => { return value.key == v["ParentID"]; });

      if (treecol.length != 0) {
        treecol[0].children.push(tree);
      }

    });

    // var remainingArr = treearr.filter(data => data.data === 0);

    var remainingArr = treearr.filter(data => data.key == 1);

    //   var finalArr = remainingArr.sort((a, b) =>{
    //     return a.label - b.label;
    // });


    console.log("REMAINING ARRAY ", remainingArr);
    console.log("REMAINING ARRAY LENGTH", remainingArr.length);

    //  Array.prototype.push.apply(allItemsFile, allItemsMain);
    // const mergedArray = [ ...allItemsFile, ...allItemsMain ];

    const unique = [];

    // mergedArray.map(x => unique.filter(a => a.key == x.key).length > 0 ? null : unique.push(x));

    // all.map(x => unique.filter(a => a.key == x.key).length > 0 ? null : unique.push(x));

    // console.log(unique);

    // tree = remainingArr;

    // this.setState({ TreeLinks: remainingArr });
    this.setState({ TreeLinks: remainingArr });
    console.log("FOLDERS", allItemsMain.length);
    // console.log("FILES", allItemsFile.length);
    // console.log("MERGED", mergedArray.length);



  }

  private getItemId() {
    var queryParms = new URLSearchParams(document.location.search.substring(1));
    var myParm = queryParms.get("folder");
    if (myParm) {
      return myParm.trim();
    }
  }



  private async getParentID(id: any) {

    var parentID = null;

    //var parentIDArray = [] ;



    await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + id + "'").get().then((results) => {
      parentID = results[0].ParentID;
      parentIDArray.push(parentID);

      console.log("Parent 1", parentID);

    });


    while (parentID != 1) {
      await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + parentID + "'").get().then((results) => {
        parentID = results[0].ParentID;
        parentIDArray.unshift(parentID);

        console.log("Parent 2", parentID);
      });
    }


    parentIDArray.push(parseInt(this.getItemId()));



    if (parentIDArray.length > 1) {
      parentIDArray.shift();
    }




    // parentIDArray.sort(function (a, b) { return a - b });
    console.log("ArrayParent", parentIDArray);


  }

  public render(): React.ReactElement<IMyGedTreeViewProps> {


    var x = this.getItemId();



    console.log("ITEM TO EXPAND", this.getItemId());

    return (

      // <div className={styles.myGedTreeView}>

      <div className="container-fluid" style={{ height: "100vh" }}>

        <div className="row" style={{ height: "100vh" }}>
          <div className="col-sm-3">
            <div id="sidebarMenu" className="sidebar">
              <div className="position-sticky">
                <div className="list-group list-group-flush mx-3 mt-4" id="tree">

                  <TreeView

                    items={this.state.TreeLinks}
                    defaultExpanded={true}
                    defaultExpandedKeys={parentIDArray}
                    // defaultExpandedKeys={[1, 9, 196, 216, 221, 224]}

                    // selectionMode={TreeViewSelectionMode.Multiple}
                    showCheckboxes={false}
                    treeItemActionsDisplayMode={TreeItemActionsDisplayMode.Buttons}
                    // onSelect={this.onSelect}
                    onExpandCollapse={this.onTreeItemExpandCollapse}
                    onRenderItem={this.renderCustomTreeItem}
                    // defaultSelectedKeys={[parseInt(this.getItemId())]}
                    defaultSelectedKeys={[parseInt(x)]}
                    expandToSelected={true}
                    defaultExpandedChildren={false}

                  />

                </div>
              </div>


            </div>

            <div className="loader-container" id="spinner">
              <div className="spinner"></div>
            </div>

          </div>

          <div className="col-sm-9">


            <form id="form_metadata">



              <div id="doc_form">
                <div className="container">
                  <div className="image">
                    <img src='https://ncaircalin.sharepoint.com/sites/TestMyGed/SiteAssets/images/flower.png' />
                    <h2 id='h2_title'>
                    </h2>
                  </div>

                </div>

                <nav aria-label="breadcrumb" id='nav_file'>
                  <ul className="breadcrumb">
                    <li className="breadcrumb-item"><a href="#" role="button" title="Mettre à jour le document" onClick={(event: React.MouseEvent<HTMLElement>) => { this.load_folders(); $("#edit_cancel_doc").css("display", "block"); $(".update_details_doc").css("display", "block"); }}><FontAwesomeIcon icon={faEdit} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                    <li className="breadcrumb-item" id='view_doc'></li>
                    <li className="breadcrumb-item"><a href="#" role="button" title="Autorisation sur le document" onClick={(event: React.MouseEvent<HTMLElement>) => { $("#access_rights_form").css("display", "block"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "none"); }}><FontAwesomeIcon icon={faLock} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                    <li className="breadcrumb-item"><a href="#" role="button" title="Télécharger" onClick={(event: React.MouseEvent<HTMLElement>) => { $("#access_rights_form").css("display", "none"); $("#subfolders_form").css("display", "block"); $("#edit_details").css("display", "none"); }}><FontAwesomeIcon icon={faDownload} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                    <li className="breadcrumb-item" id='delete_document'></li>
                    <li className="breadcrumb-item"><a href="#" role="button" title="Notifier" ><FontAwesomeIcon icon={faBell} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                  </ul>
                </nav>



                <div id="doc_details">

                  <div className="row">
                    {/* <div className="col-6">
                      <Label>Title
                        <input type="text" className="form-control" id="input_title" />
                      </Label>
                    </div> */}
                    <div className="col-6">
                      <Label>Document Number
                        <input type="text" id='input_number' className='form-control' />
                      </Label>
                    </div>
                    <div className="col-6">
                      <Label>Dossier
                        <input type="text" className="form-control" id="input_type_doc" list='folders' />

                        <datalist id="folders">
                          <select id="select_folders"></select>
                        </datalist>
                      </Label>
                    </div>

                  </div>

                  <div className="row">
                    <div className="col-3">
                      <Label>
                        Revision
                        <input type="text" id='input_revision' className='form-control' />
                      </Label>
                    </div>
                    <div className="col-3">
                      <Label>
                        Status
                        <input type="text" id='input_status' className='form-control' />
                      </Label>
                    </div>
                    <div className="col-3">
                      <Label>
                        Owner
                        <input type="text" id='input_owner' className='form-control' />
                      </Label>
                    </div>
                    <div className="col-3">
                      <Label>
                        Active Date
                        <input type="text" id='input_activeDate' className='form-control' />
                      </Label>
                    </div>
                  </div>

                  <div className="row">
                    <div className="col-6">
                      <Label>
                        Filename
                        <input type="text" id='input_filename' className='form-control' />
                      </Label>
                    </div>
                    <div className="col-6">
                      <Label>
                        Author
                        <input type="text" id='input_author' className='form-control' />
                      </Label>
                    </div>

                  </div>

                  <div className="row">
                    <div className="col-5">
                      <Label>
                        Description
                        <textarea id='input_description' className='form-control' rows={2} />
                      </Label>
                    </div>
                    <div className="col-4">
                      <Label>
                        Keywords
                        <textarea id='input_keywords' className='form-control' rows={2} />
                      </Label>
                    </div>
                    <div className="col-3">
                      <Label>
                        Review Date
                        <input type="text" id='input_reviewDate' className='form-control' />
                      </Label>
                    </div>
                  </div>

                  <div className="row">
                    <div className="col-8">

                    </div>
                    <div className="col-2" id='btn_update_doc'>


                    </div>

                    <div className="col-2">
                      <button type="button" className="btn btn-primary" id='edit_cancel_doc'>Cancel</button>
                    </div>

                  </div>

                </div>

                <div id="doc_permission"></div>
                <div id="doc_notif"></div>


              </div>

              <div id="access_form">
                <div className="container">
                  <div className="image">
                    <img src='https://ncaircalin.sharepoint.com/sites/TestMyGed/SiteAssets/images/flower.png' />
                    <h2 id='h2_folderName'>
                    </h2>
                  </div>

                </div>


                <nav aria-label="breadcrumb" id='nav'>
                  <ul className="breadcrumb" id="folder_nav">
                    <li className="breadcrumb-item"><a href="#" title="Mettre à jour le dossier" role="button" onClick={(event: React.MouseEvent<HTMLElement>) => { this.load_folders(); $("#access_rights_form").css("display", "none"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "block"); $("#doc_details_add").css("display", "none"); }}><FontAwesomeIcon icon={faEdit} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                    <li className="breadcrumb-item"><a href="#" title="Créer un document" role="button" id='add_document' onClick={(event: React.MouseEvent<HTMLElement>) => { $("#access_rights_form").css("display", "none"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "none"); $("#doc_details_add").css("display", "block"); }}><FontAwesomeIcon icon={faFile} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                    <li className="breadcrumb-item"><a href="#" title="Autorisation sur le dossier" role="button" onClick={(event: React.MouseEvent<HTMLElement>) => { $("#access_rights_form").css("display", "block"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "none"); $("#doc_details_add").css("display", "none"); }}><FontAwesomeIcon icon={faLock} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                    <li className="breadcrumb-item"><a href="#" title="Ajouter des sous-dossiers" role="button" onClick={(event: React.MouseEvent<HTMLElement>) => { $("#access_rights_form").css("display", "none"); $("#subfolders_form").css("display", "block"); $("#edit_details").css("display", "none"); $("#doc_details_add").css("display", "none"); }}><FontAwesomeIcon icon={faFolderPlus} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                    <li className="breadcrumb-item" id='bouton_delete'></li>

                    {/* <li className="breadcrumb-item" id='bouton_delete'><a href="#" title="Supprimer" role="button" id='delete_folder'><FontAwesomeIcon icon={faTrashCan} className="fa-icon fa-2x"></FontAwesomeIcon></a></li> */}
                    <li className="breadcrumb-item"><a href="#" title="Notifier" role="button" ><FontAwesomeIcon icon={faBell} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                  </ul>
                </nav>


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

                    <div className="col-6">
                      <Label>Folder Order
                        <input type="text" className="form-control" id="folder_order" />
                      </Label>
                    </div>


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
                      <Label>Group name
                        <input type="text" className="form-control" id="group_name" />
                      </Label>
                    </div>
                    <div className="col-3">
                      <Label>Permission Type
                        <select className='form-select' name="permissions" id="permissions">
                          <option value="none">NONE</option>
                          <option value="read">READ</option>
                          <option value="read_write">READ_WRITE</option>
                          <option value="all">ALL</option>
                        </select>
                      </Label>
                    </div>
                    <div className="col-3" id="add_btn_group">
                    </div>
                  </div>
                </div>

                <div id="subfolders_form">
                  <div className="row">
                    <div className="col-8">
                      <Label>Folder name
                        <input type="text" className="form-control" id="folder_name" />
                      </Label>
                    </div>

                    <div className="col-3" id="add_btn_subFolder">

                    </div>
                  </div>

                </div>

                <div id="doc_details_add">

                  <div className="row">
                    {/* <div className="col-6">
            <Label>Title
                <input type="text" className="form-control" id="input_title_add" />
            </Label>
        </div> */}
                    <div className="col-6">
                      <Label>Nom du fichier
                        <input type="text" id='input_doc_number_add' className='form-control' />
                      </Label>
                    </div>


                    <div className="col-6">
                      <Label>Fichier
                        <input type="file" name="file" id="file_ammendment" className="form-control-file mb-3" />
                      </Label>

                      {/* <Form.Group controlId="formFile" className="mb-3">
                        <Form.Label>Fichier</Form.Label>
                        <Form.Control type="file" />
                      </Form.Group> */}

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
                        <input type="text" id='input_filename_add' className='form-control' />
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


  private async addSubfolders(item: ITreeItem) {

    console.log("ID", item.id);
  }

  private async onTreeItemSelect(items: ITreeItem[]) {

    items.forEach((item: any) => {
      $("#h2_folderName").text(item.label);
    });

  }

  private async folderActions(id: any) {



  }

  private onSelect(items: ITreeItem[]) {



  }

  private async onTreeItemExpandCollapse(item: ITreeItem, isExpanded: boolean) {
    console.log((isExpanded ? "Item expanded: " : "Item collapsed: ") + item.label);
    console.log(item.key);
    console.log(item.key);

    if (isExpanded) {


    }

    $("#text").text(item.label);
  }

  private renderCustomTreeItem(item: ITreeItem): JSX.Element {


    return (
      <span

        // onClick={(event: React.MouseEvent<HTMLElement>)
        onClick={async (event: React.MouseEvent<HTMLInputElement>) => {


          {
            console.log("DATA value", item.label);


            if (item.data == 1 || item.data == 0) {




              var fileName = "";
              var content = null;

              var titleFolder = "";

              const allItemsFolder: any[] = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID,Title,revision,IsFolder,description").filter("FolderID eq '" + item.parentID + "'").getAll();

              allItemsFolder.forEach((x) => {

                titleFolder = x.Title;

              });


              // window.history.pushState("object or string", "Title", "https://ncaircalin.sharepoint.com/sites/TestMyGed/SitePages/Home.aspx?folder=6" + item.key);


              console.log("TITLE", titleFolder);

              $("#h2_folderName").text(item.label);

              $("#folder_name1").val(item.label);
              $("#folder_desc").val(item.description);
              $("#parent_folder").val(item.parentID + "_" + titleFolder);
              // $("#h2_folderName").text(item.label);


              $("#access_form").css("display", "block");

              $("#doc_form").css("display", "none");


              //delete dossier

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

              //update dossier


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


              //upload document

              var add_doc_container: Element = document.getElementById("add_document_btn");

              let add_btn_document: string = `
              <button type="button" class="btn btn-primary add_doc" id=${item.id}_add_doc>Sauvegarder</button>
              `;

              add_doc_container.innerHTML = add_btn_document;


              const btn_add_doc = document.getElementById(item.id + '_add_doc');

              await btn_add_doc?.addEventListener('click', async () => {


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


              //add_subfolder

              var add_subfolder_container: Element = document.getElementById("add_btn_subFolder");

              let add_btn_subfolder: string = `
              <button type="button" class="btn btn-primary add_subfolder mb-2 " id=${item.id}_add_btn_subfolder>Add subfolder</button>
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


              //add group permission

              var add_group_permission_container: Element = document.getElementById("add_btn_subFolder");

              let add_btn_group_permission: string = `
              <button type="button" class="btn btn-primary add_group mb-2" id=${item.id}_add_group>Add group</button>
              `;

              add_group_permission_container.innerHTML = add_btn_group_permission;


              const btn_add_group = document.getElementById(item.id + '_add_group');

              await btn_add_group?.addEventListener('click', async () => {

                try {
                  console.log("KEY", item.key);

                  await sp.web.lists.getByTitle("AccessRights").items.add({
                    Title: item.label,
                    groupName: $("#group_name").val(),
                    permission: $("#permissions option:selected").val(),
                    FolderIDId: item.key
                  })
                    .then(() => {
                      console.log("Autorisation ajoutée à ce dossier avec succès.")
                    })
                    .then(() => {
                      window.location.reload();
                    });
                }

                catch (e) {
                  alert("Erreur: " + e.message);
                }


              });





              $('#file_ammendment').on('change', () => {
                const input = document.getElementById('file_ammendment') as HTMLInputElement | null;


                var file = input.files[0];
                var reader = new FileReader();

                reader.onload = ((file1) => {
                  return (e) => {
                    console.log(file1.name);

                    fileName = file1.name,
                      content = e.target.result

                  };
                })(file);

                reader.readAsArrayBuffer(file);
              });


            }

            else {
              $("#doc_details").css("display", "block");

              $("#nav_file").css("display", "block");

              $("#access_form").css("display", "none");

              $("#doc_form").css("display", "block");

              var urlFile = '';

              var titleFolder = "";

              const allItemsFolder: any[] = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID,Title,revision,IsFolder,description").filter("FolderID eq '" + item.parentID + "'").getAll();

              allItemsFolder.forEach((x) => {

                titleFolder = x.Title;

              });

              await sp.web.lists.getByTitle('Documents').items
                .select('Id', 'Title')
                .get()
                .then(response => {
                  response
                    .forEach(x => {
                      var _Item = sp.web.lists.getByTitle("Documents")
                        .items
                        .getById(parseInt(item.id));

                      console.log("ITEMS", _Item);


                      _Item.attachmentFiles
                        .select('FileName', 'ServerRelativeUrl')
                        .get()
                        .then(responseAttachments => {
                          responseAttachments
                            .forEach(attachmentItem => {
                              // result += item.Title                       + "<br/>" +
                              // item.EncodedAbsUrl                    + "<br/>" + 
                              // attachmentItem.FileName          + "<br/>" + 
                              urlFile = attachmentItem.ServerRelativeUrl;

                              // $("#input_title").val(x.T);
                              // $("#input_url").val(item_detail.FileUrl);
                              // window.open(`${attachmentItem.ServerRelativeUrl}`, '_blank');

                            });
                        });

                    });
                });



              $('#edit_cancel_doc').one('click', async function () {

                $(".update_details_doc").css("display", "none");

              });

              //update doc


              var update_doc_container: Element = document.getElementById("btn_update_doc");

              let update_btn_doc: string = `<button type="button" class="btn btn-primary update_details_doc" id='${item.id}_update_details_doc'>Edit Details</button>`;

              update_doc_container.innerHTML = update_btn_doc;


              const btn_edit_doc = document.getElementById(item.id + '_update_details_doc');

              await btn_edit_doc?.addEventListener('click', async () => {

                let text = $("#input_type_doc").val();
                const myArray = text.toString().split("_");
                let parentId = myArray[0];


                if (confirm(`Etes-vous sûr de vouloir mettre à jour les détails de ${item.label} ?`)) {

                  try {

                    const i = await await sp.web.lists.getByTitle('Documents').items.getById(parseInt(item.id)).update({
                      Title: $("#input_number").val(),
                      description: $("#input_description").val(),
                      keywords: $("#input_keywords").val(),
                      doc_number: $("#input_number").val(),
                      revision: $("#input_revision").val(),
                      ParentID: parseInt(parentId)

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


              });


              //delete doc
              var delete_doc: Element = document.getElementById("delete_document");


              let nav_html_delete_doc: string = '';


              // console.log("ONSELECT", item.label);

              nav_html_delete_doc = `
                              <a href="#" title="Supprimer" 
                              role="button" id='${item.id}_deleteDoc'>
                          <svg aria-hidden="true" focusable="false" data-prefix="far" data-icon="trash-can" 
                          class="svg-inline--fa fa-trash-can fa-icon fa-2x" role="img" xmlns="http://www.w3.org/2000/svg" 
                          viewBox="0 0 448 512">
                          <path fill="currentColor" d="M160 400C160 408.8 152.8 416 144 416C135.2 416 128 408.8 128 400V192C128 183.2 135.2 176 144 176C152.8 176 160 183.2 160 192V400zM240 400C240 408.8 232.8 416 224 416C215.2 416 208 408.8 208 400V192C208 183.2 215.2 176 224 176C232.8 176 240 183.2 240 192V400zM320 400C320 408.8 312.8 416 304 416C295.2 416 288 408.8 288 400V192C288 183.2 295.2 176 304 176C312.8 176 320 183.2 320 192V400zM317.5 24.94L354.2 80H424C437.3 80 448 90.75 448 104C448 117.3 437.3 128 424 128H416V432C416 476.2 380.2 512 336 512H112C67.82 512 32 476.2 32 432V128H24C10.75 128 0 117.3 0 104C0 90.75 10.75 80 24 80H93.82L130.5 24.94C140.9 9.357 158.4 0 177.1 0H270.9C289.6 0 307.1 9.358 317.5 24.94H317.5zM151.5 80H296.5L277.5 51.56C276 49.34 273.5 48 270.9 48H177.1C174.5 48 171.1 49.34 170.5 51.56L151.5 80zM80 432C80 449.7 94.33 464 112 464H336C353.7 464 368 449.7 368 432V128H80V432z"></path></svg> 
                              </a>
                              
                              `;

              delete_doc.innerHTML = nav_html_delete_doc;

              const btn_delete_doc = document.getElementById(item.id + '_deleteDoc');


              await btn_delete_doc?.addEventListener('click', async () => {
                if (confirm(`Êtes-vous sûr de vouloir supprimer ${item.label} ?`)) {

                  try {
                    var res = await sp.web.lists.getByTitle('Documents').items.getById(parseInt(item.id)).delete()
                      .then(() => {
                        alert("Document supprimé avec succès.");
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



              //view document

              var view_doc: Element = document.getElementById("view_doc");


              let nav_html_view_doc: string = '';


              // console.log("ONSELECT", item.label);

              nav_html_view_doc = `
              <a href="#" role="button" title="Voir le document" id="${item.id}_view_doc">
              <svg aria-hidden="true" focusable="false" data-prefix="far" 
              data-icon="eye" class="svg-inline--fa fa-eye fa-icon fa-2x" 
              role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512">
              <path fill="currentColor" d="M160 256C160 185.3 217.3 128 288 128C358.7 128 416 185.3 416 256C416 326.7 358.7 384 288 384C217.3 384 160 326.7 160 256zM288 336C332.2 336 368 300.2 368 256C368 211.8 332.2 176 288 176C287.3 176 286.7 176 285.1 176C287.3 181.1 288 186.5 288 192C288 227.3 259.3 256 224 256C218.5 256 213.1 255.3 208 253.1C208 254.7 208 255.3 208 255.1C208 300.2 243.8 336 288 336L288 336zM95.42 112.6C142.5 68.84 207.2 32 288 32C368.8 32 433.5 68.84 480.6 112.6C527.4 156 558.7 207.1 573.5 243.7C576.8 251.6 576.8 260.4 573.5 268.3C558.7 304 527.4 355.1 480.6 399.4C433.5 443.2 368.8 480 288 480C207.2 480 142.5 443.2 95.42 399.4C48.62 355.1 17.34 304 2.461 268.3C-.8205 260.4-.8205 251.6 2.461 243.7C17.34 207.1 48.62 156 95.42 112.6V112.6zM288 80C222.8 80 169.2 109.6 128.1 147.7C89.6 183.5 63.02 225.1 49.44 256C63.02 286 89.6 328.5 128.1 364.3C169.2 402.4 222.8 432 288 432C353.2 432 406.8 402.4 447.9 364.3C486.4 328.5 512.1 286 526.6 256C512.1 225.1 486.4 183.5 447.9 147.7C406.8 109.6 353.2 80 288 80V80z">
              </path></svg>
              
              </a>
                              
                              `;

              view_doc.innerHTML = nav_html_view_doc;

              const btn_view_doc = document.getElementById(item.id + '_view_doc');

              await btn_view_doc?.addEventListener('click', async () => {
                window.open(`${urlFile}`, '_blank');
              });





              //  var item1: any = sp.web.lists.getByTitle('Documents').items.getById(parseInt(item.id));
              const item1: any = await sp.web.lists.getByTitle("Documents").items.getById(parseInt(item.id))();
              console.log(item1);

              console.log(item1);

              Object.keys(item1).forEach((key) => {

                // $("#input_title").val(item1.Title);
                $("#input_type_doc").val(item.parentID+"_"+titleFolder);
                $("#input_number").val(item1.doc_number);
                $("#input_revision").val(item1.revision);
                $("#input_status").val(item1.status);
                $("#input_owner").val(item1.owner);
                $("#input_activeDate").val(item1.active_date);
                $("#input_filename").val(item1.filename);
                $("#input_author").val(item1.author);
                // $("#input_reviewDate").val(item1.);
                $("#input_keywords").val(item1.keywords);
                $("#input_description").val(item1.description);
                $("#h2_title").text(item1.Title);

              });



            }
          }

        }

        }
      >



        {
          <FontAwesomeIcon icon={item.icon} className="fa-icon"></FontAwesomeIcon>
        }

        &nbsp;
        {item.label}

      </span>
    );


  }

}


