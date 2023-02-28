import styles from './MyGedTreeView.module.scss';
import { MSGraphClient } from '@microsoft/sp-http';
import { IMyGedTreeViewProps, IMyGedTreeViewState } from './IMyGedTreeView';
import { TreeView, ITreeItem, TreeViewSelectionMode, TreeItemActionsDisplayMode } from "@pnp/spfx-controls-react/lib/TreeView";
import 'bootstrap/dist/css/bootstrap.min.css';
import $ from 'jquery';
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
import { faFile, faLock, faFolderPlus, faDownload, faMagnifyingGlass, faDeleteLeft } from '@fortawesome/free-solid-svg-icons'
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



// import {
//   SPHttpClient,
//   SPHttpClientResponse
// } from '@microsoft/sp-http';



// var parentIDArray: number[];

// var parentIDArray = new Array();

// var parentIDArray = [];

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

var remainingArr: any = [];
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
      // spfxContext: this.context

      //props.context
    });

    // this.context.msGraphClientFactory
    // .getClient()
    // .then((client: MSGraphClient): void => {
    //     this.graphClient = client;
    //     resolve();
    // }, err => reject(err));


    // var x = this.getItemId();

    //this.getParentID(x);

    // const sp = spfi().using(SPFx(this.props.context));
    this.state = {
      TreeLinks: [],
      parentIDArray: [],
    };


    var x = this.getItemId();

    this._getLinks2(sp);

    // this._getLinks3(sp);  //sa pu tester doc library

    // this._getLinks(sp);
    // this.render();

    this.loadDocsFromFolders(x);

  }



  // protected onInit(): Promise<void> {
  //   return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
  //     // this.user = this.context.pageContext.user;
  //     sp.setup({
  //       spfxContext: this.context
  //     });

  //     this.context.msGraphClientFactory
  //       .getClient()
  //       .then((client: MSGraphClient): void => {
  //         this.graphClient = client;
  //         resolve();
  //       }, err => reject(err));
  //   });
  // }


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
    // const allItemsFile: any[] = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID,Title,revision,IsFolder,description").filter("IsFolder eq '" + value2 + "'").getAll();


    //const allItemsMain_sorted: any[] = allItemsMain.sort((a, b) => { return a.Title - b.Title });
    // const allItemsMain_sorted: any[] = allItemsMain.sort();


    console.log("ARRRANGED ARRAY: " + allItemsMain.length);

    // const allItemsMain_sorted: any[] = allItemsMain.sort((a, b) => a.FolderID - b.FolderID);
    var x = 0;

    await Promise.all(allItemsMain.map(async (v) => {


      // allItemsMain.forEach(v => {


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

    }));


    keysMissing = allKeys
      .filter(x => !keysPresent.includes(x))
      .concat(keysPresent.filter(x => !allKeys.includes(x)));

    keysMissing.forEach(async (v) => {

      await Promise.all(allItemsMain.map(async (x) => {

        // allItemsMain.forEach(x => {

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
      }));

    });


    // console.log("KEYS MISSING LENGTH", keysMissing.length);
    // console.log("KEYS MISSING", keysMissing);
    // console.log("KEYS PRESENT LENGTH", keysPresent.length);
    // console.log("KEYS MISSING LENGTH", 763 - keysPresent.length);
    // console.log("COUNTER", counter);
    // console.log("COUNTERSUB", counterSUB);
    // console.log("TREE ARRAY LENGTH", treearr.length);

    // allItemsFile.forEach((v) => {
    //   // console.log("We have a file here.");


    //   //v["ServerRedirectedEmbedUrl"] ==> v["FileUrl"]
    //   const iconName: IconProp = faFile;

    //   const tree: ITreeItem = {
    //     // v.id ==> v.FolderID
    //     id: v["ID"],
    //     //    key: v["FolderID"],
    //     key: v["FolderID"],
    //     label: v["Title"] + "-" + v["revision"],
    //     icon: faFile,
    //     data: v["FileUrl"],
    //     revision: v['revision'],
    //     file: "Yes",
    //     description: v["description"],
    //     parentID: v["ParentID"]

    //   };


    //   // ParentID ==> FolderID

    //   var treecol: Array<ITreeItem> = treearr.filter((value) => { return value.file === "No" && value.key == v["ParentID"]; });

    //   if (treecol.length != 0) {
    //     treecol[0].children.push(tree);
    //   }

    // });

    // var remainingArr = treearr.filter(data => data.data === 0);

    //mn tir var ldns
    remainingArr = treearr.filter(data => data.key == 1);



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

  private async getParentArray(id: any, arrayParent: any) {


    var parentID = null;

    //var parentIDArray = [] ;

    await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + id + "'").get().then((results) => {
      parentID = results[0].ParentID;
      arrayParent.push(parseInt(parentID));

      console.log("Parent 1", parentID);

    });


    while (parentID != 1) {
      await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + parentID + "'").get().then((results) => {
        parentID = results[0].ParentID;
        arrayParent.unshift(parseInt(parentID));

        console.log("Parent 2", parentID);
      });
    }


    arrayParent.push(parseInt(this.getItemId()));



    if (arrayParent.length > 1) {
      arrayParent.shift();
    }

    return arrayParent;


  }

  public async componentDidMount() {
    var x = this.getItemId();

    const parentIDs = await this.getParentID(x);
    this.setState({parentIDArray:parentIDs});

    console.log("INSIDE THE DID MOUNT", parentIDs, this.state.parentIDArray)

    // this.render();

  }


  public async componentWillMount() {
    var x = this.getItemId();

    const parentIDs = await this.getParentID(x);
    this.setState({parentIDArray:parentIDs});

    console.log("INSIDE THE WILL MOUNT", parentIDs, this.state.parentIDArray)

    // this.render();

  }


  private async getParentID(id: any) {

    var parentID = null;

    var parentIDArray = [] ;

    await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + id + "'").get().then((results) => {
      parentID = results[0].ParentID;

      // this.setState({ parentIDArray: [...this.state.parentIDArray, parentID] });
      parentIDArray.push(parentID);

      console.log("Parent 1", parentID);

    });


    while (parentID != 1) {
      await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + parentID + "'").get().then((results) => {
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

    var y = [];

    x = this.getItemId();

    // this.getParentID(x);


    console.log("TEST PARENT ARRAY", y);

    console.log("ITEM TO EXPAND", this.getItemId());

    // (async () => {
    //   y = await this.getParentArray(x, y);
    // })();


    // useAsyncEffect(async () => {
    //   y = await this.getParentArray(x, y);
    // });

    // var x = this.getItemId();
    // (async () => {
    //   await this.getParentID(x);
    //   this.render();

    // })();
    console.log("BEFORE RENDER", this.state.parentIDArray);

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
                    
                    // defaultExpandedKeys={[212, 213, 243, 244, 248, 249]}
                    // defaultExpandedKeys={y}

                    // selectionMode={TreeViewSelectionMode.Multiple}
                    showCheckboxes={false}
                    treeItemActionsDisplayMode={TreeItemActionsDisplayMode.Buttons}
                    onSelect={this.onSelect}

                    onExpandCollapse={this.onTreeItemExpandCollapse}
                    onRenderItem={this.renderCustomTreeItem}
                    // defaultSelectedKeys={[parseInt(this.getItemId())]}
                     defaultSelectedKeys={[this.state.parentIDArray[0], parseInt(x)]}

                     defaultExpandedKeys={this.state.parentIDArray}
                    // defaultSelectedKeys={this.state.parentIDArray}
                    // defaultSelectedKeys={y}

                    expandToSelected={true}
                    defaultExpandedChildren={false}

                  />

                </div>
              </div>

            </div>
          </div>

          <div className="col-sm-9">

            <div id="loader"></div>

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
                    <li className="breadcrumb-item"><a href="#" role="button" title="Mettre à jour le document" onClick={async (event: React.MouseEvent<HTMLElement>) => {
                      await this.load_folders(); $("#access_form").css("display", "none");
                      $("#doc_form").css("display", "block");
                      $("#doc_details").css("display", "block");
                      $("#table_documents").css("display", "none");
                      $("#nav_file").css("display", "block");
                      $("#notifications_doc_form").css("display", "none");
                    }}><FontAwesomeIcon icon={faEdit} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                    <li className="breadcrumb-item" id='view_doc'></li>
                    {/* <li className="breadcrumb-item"><a href="#" role="button" title="Autorisation sur le document" onClick={async (event: React.MouseEvent<HTMLElement>) => { await this.getSiteUsers(); await this.getSiteGroups(); $("#doc_form_access_rights").css("display", "block"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "none"); }}><FontAwesomeIcon icon={faLock} className="fa-icon fa-2x"></FontAwesomeIcon></a></li> */}

                    <li className="breadcrumb-item"><a href="#" title="Autorisation sur le document" role="button" onClick={async (event: React.MouseEvent<HTMLElement>) => { await this.getSiteUsers(); this.getSiteGroups(); $("#table_documents").css("display", "none"); $("#access_rights_form").css("display", "block"); $("#alert_0_doc").css("display", "none"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "none"); $("#doc_details_add").css("display", "none"); $("#doc_details").css("display", "none"); $("#table_version_doc").css("display", "none"); $(".dossier_headers").css("display", "none"); $("#access_form").css("display", "block"); $("#notifications_doc_form").css("display", "none"); }}><FontAwesomeIcon icon={faLock} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>

                    <li className="breadcrumb-item" id='download_doc'></li>
                    <li className="breadcrumb-item" id='delete_document'></li>
                    <li className="breadcrumb-item"><a href="#" role="button" title="Notifier" onClick={async (event: React.MouseEvent<HTMLElement>) => { await this.getSiteUsers(); this.getSiteGroups(); $("#table_documents").css("display", "none"); $("#access_rights_form").css("display", "block"); $("#alert_0_doc").css("display", "none"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "none"); $("#doc_details_add").css("display", "none"); $("#doc_details").css("display", "none"); $("#table_version_doc").css("display", "none"); $(".dossier_headers").css("display", "none"); $("#access_form").css("display", "none"); $("#notifications_doc_form").css("display", "block"); }} ><FontAwesomeIcon icon={faBell} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>


                  </ul>

                </nav>


                {<div id="doc_details">

                  <legend>Détails</legend>

                  <div className="row">

                    <div className="col-6">
                      <Label>Nom du fichier
                        <input type="text" id='input_number' className='form-control' disabled />
                      </Label>
                    </div>


                    <div className="col-6">
                      <Label>Dossier
                        <input type="text" className="form-control" id="input_type_doc" list='folders' disabled />

                        <datalist id="folders">
                          <select id="select_folders"></select>
                        </datalist>
                      </Label>
                    </div>

                    {/* 
                    <div className="col-6">
                      <Label>Dossier
                        <input type="text" className="form-control form-control-lg" id="input_type_doc" list='folders' disabled />

                        <datalist id="folders">
                          <select id="select_folders"></select>
                        </datalist>
                      </Label>
                    </div> */}

                  </div>

                  <div className="row">
                    <div className="col-8">
                      <Label>
                        Description
                        <textarea id='input_description' className='form-control' rows={2} />
                      </Label>
                    </div>
                    <div className="col-4">
                      <Label>
                        Mots-clés
                        <textarea id='input_keywords' className='form-control' rows={2} />
                      </Label>
                    </div>

                  </div>

                  <div className='row'>
                    <div className="col-4">
                      <Label>
                        Review Date
                        <input type="text" id='input_reviewDate' className='form-control' disabled />
                      </Label>
                    </div>

                    <div className="col-4">
                      <Label>
                        Owner
                        <input type="text" id='created_by' className='form-control' disabled />
                      </Label>
                    </div>

                    <div className="col-4">
                      <Label>
                        Date de création
                        <input type="text" id='creation_date' className='form-control' disabled />
                      </Label>
                    </div>


                  </div>

                  <legend>Détails de dernière mise à jour</legend>

                  <div className="row">

                    <div className="col-8">
                      <Label>
                        Revision
                        <input type="text" id='input_revision' className='form-control' />
                      </Label>
                    </div>

                    <div className="col-4">
                      <Label>
                        Status
                        <input type="text" id='input_status' className='form-control' hidden />
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


                    <div className="col-4">
                      <Label>Fichier
                        <input type="file" name="file" id="file_ammendment_update" className="form-control" />
                      </Label>

                    </div>

                    <div className="col-4">
                      <Label>
                        Filename
                        <input type="text" id='input_filename' className='form-control' disabled />
                      </Label>
                    </div>

                    <div className="col-4">
                      <Label>
                        Author
                        <input type="text" id='input_author' className='form-control' />
                      </Label>
                    </div>

                  </div>

                  <div className='row'>

                    <div className="col-4">
                      <Label>
                        Updated
                        <input type="text" id='updated_by' className='form-control' disabled />
                      </Label>
                    </div>

                    <div className="col-4">
                      <Label>
                        Date
                        <input type="text" id='updated_time' className='form-control' disabled />
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

                </div>}

                {/* <div id="doc_form_access_rights">


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

                </div> */}


                <div id="doc_permission"></div>


                <div id="notifications_doc_form">


                  <h3>Notification</h3>
                  <div className="row">

                    <div className="col-6">
                      <Label>Ajouter une notification utilisateur :

                        <input type="text" className="form-control" id="users_name_notif" list='users' />

                        <datalist id="users">
                          <select id="select_users"></select>
                        </datalist>
                      </Label>
                    </div>



                    <div className="col-3" id="add_notif_btn_user_doc">
                    </div>
                  </div>

                  <div className="row">


                    <div className="col-6">
                      <Label>Ajouter une notification de groupe :
                        <input type="text" className="form-control" id="group_name" list='groups' />

                        <datalist id="groups">
                          <select id="select_groups"></select>
                        </datalist>
                      </Label>
                    </div>


                    <div className="col-3" id="add_notif_btn_group_doc">
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

                <div id='table_version_doc'>

                  <div id="spListDocumentsVersions">

                    <table id='tbl_documents_versions' className='table table-striped'>
                      <thead>
                        <tr>
                          <th className="text-left">ID</th>
                          <th className="text-left">Nom du document</th>
                          <th className="text-left" >Version</th>

                          <th className="text-right" >Actions</th>
                        </tr>
                      </thead>
                      <tbody id="tbl_documents_versions_bdy">



                      </tbody>
                    </table>
                  </div>


                </div>

                <div className="modal fade right" id="exampleModalPreview" tabIndex={-1} role="dialog"
                  aria-labelledby="exampleModalPreviewLabel" aria-hidden="true" data-backdrop="false">
                  <div className="modal-dialog-full-width modal-dialog momodel modal-fluid" role="document">
                    <div className="modal-content-full-width modal-content">
                      <div className=" modal-header-full-width   modal-header text-center">
                        <h5 className="modal-title w-100" id="exampleModalPreviewLabel">Contract</h5>
                        <button type="button" className="close " data-dismiss="modal" aria-label="Close">

                        </button>
                      </div>
                      <div className="modal-body">
                        <div id="iframe_word"></div>

                      </div>
                      <div className="modal-footer-full-width  modal-footer">
                        <button type="button" className="btn btn-danger btn-md btn-rounded"
                          data-dismiss="modal">Close</button>

                      </div>
                    </div>
                  </div>
                </div>


              </div>

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
                          <th className="text-left">ID</th>
                          <th className="text-left">Nom du document</th>
                          <th className="text-left" >Description</th>

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
                        <Label>Nom du fichier
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

  private onSelect(items: ITreeItem[]) {



    items.forEach(async (item) => {

      $("#h2_folderName").text(item.label);


    });



  }

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


  private renderCustomTreeItem(item: ITreeItem): JSX.Element {


    return (
      <span

        onClick={async (event: React.MouseEvent<HTMLInputElement>) => {
          // onChange={async (event: React.MouseEvent<HTMLInputElement>) => {
          // document.getElementById("loader").style.display = "block";

          //loader
          //   myVar = setTimeout((document.getElementById("form_metadata").style.display = "block", document.getElementById("loader").style.display = "none"), 3000)

          const groupTitle = [];
          let groups: any = await sp.web.currentUser.groups();

          usersGroups = groups;

          console.log("USERS GROUPS", usersGroups);

          usersGroups.forEach((item) => {

            groupTitle.push(item.Title);
          });


          console.log("DANS NUVO GROUP ARRAY", groupTitle);


          // if (groupTitle.includes("myGed Visitors")) {
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
            // $("#table_documents").css("display", "block");
            $("#h2_folderName").text(item.label);
          }

          $("#h2_folderName").text(item.label);



          //render table
          {

            var response_doc = null;
            var response_distinc = [];
            var html_document: string = ``;
            var value1 = "FALSE";

            var pdfName = '';


            var document_container: Element = document.getElementById("tbl_documents_bdy");
            document_container.innerHTML = "";

            //  await sp.web.lists.getByTitle('Documents').items

            const all_documents: any[] = await sp.web.lists.getByTitle('Documents').items
              .select("Id,ParentID,FolderID,Title,revision,IsFolder,description")
              .filter("ParentID eq '" + item.key + "' and IsFolder eq '" + value1 + "'")
              .get();

            response_doc = all_documents;

            //  var result = response_doc.filter((value, index, array) => array.lastIndexOf(value) === index);

            var result = response_doc.filter((obj, pos, arr) => {
              return arr.map(mapObj =>
                mapObj.Title).lastIndexOf(obj.Title) == pos;
            });



            // var result = response_doc.reduce((acc, obj) => {
            //   let last = acc.find(el => el.Title === obj.Title);
            //   if (!last || parseInt(obj.Id) > parseInt(last.Id)) {
            //     acc.push(obj);
            //     response_distinc.push(obj);
            //   }
            //   return acc;
            // }, []);




            // return arr.reduce((maxIndex, mapObj, index) =>
            //   (mapObj.Title === obj.Title && mapObj.Id > obj.Id) ? index : maxIndex, -1) === pos;



            console.log("ALL", response_doc);



            // var result = [...response_doc.reduce((r, o) =>
            //   (!r.has(o.Title) || r.get(o.Title).length < o.Id) ? r.set(o.Title, o) : r
            //   , new Map()).values()];

            //   console.log(result);



            console.log("RESULT DISTINCT", result);
            console.log("RESULT DISTINCT ARRAY LOT LA", response_distinc);


            // console.log(response_doc);

            if (result.length > 0) {


              html_document = ``;
              $("#alert_0_doc").css("display", "none");
              $("#table_documents").css("display", "block");

              // $("#table_documents").css("display", "block");


              await result.forEach(async (element) => {



                var urlFile = '';
                html_document += `
                <tr>
                <td class="text-left">${element.Id}</td>

                <td class="text-left">${element.Title}</td>

                <td class="text-left"> 
                ${element.description}          
                </td>

                
                <td>
                <a href="#" title="Mettre à jour le document" role="button" id="${element.Id}_view_doc_details" class="btn_view_doc_details">
                <svg aria-hidden="true" focusable="false" data-prefix="far" 
                data-icon="pen-to-square" class="svg-inline--fa fa-pen-to-square fa-icon fa-2x" 
                role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">
                <path fill="currentColor" d="M373.1 24.97C401.2-3.147 446.8-3.147 474.9 24.97L487 37.09C515.1 65.21 515.1 110.8 487 138.9L289.8 336.2C281.1 344.8 270.4 351.1 258.6 354.5L158.6 383.1C150.2 385.5 141.2 383.1 135 376.1C128.9 370.8 126.5 361.8 128.9 353.4L157.5 253.4C160.9 241.6 167.2 230.9 175.8 222.2L373.1 24.97zM440.1 58.91C431.6 49.54 416.4 49.54 407 58.91L377.9 88L424 134.1L453.1 104.1C462.5 95.6 462.5 80.4 453.1 71.03L440.1 58.91zM203.7 266.6L186.9 325.1L245.4 308.3C249.4 307.2 252.9 305.1 255.8 302.2L390.1 168L344 121.9L209.8 256.2C206.9 259.1 204.8 262.6 203.7 266.6zM200 64C213.3 64 224 74.75 224 88C224 101.3 213.3 112 200 112H88C65.91 112 48 129.9 48 152V424C48 446.1 65.91 464 88 464H360C382.1 464 400 446.1 400 424V312C400 298.7 410.7 288 424 288C437.3 288 448 298.7 448 312V424C448 472.6 408.6 512 360 512H88C39.4 512 0 472.6 0 424V152C0 103.4 39.4 64 88 64H200z"></path></svg></a>



               <a href="#"  title="Voir le document" id="${element.Id}_view_doc"  class="btn_view_doc" style="padding-left: inherit;">
               <svg aria-hidden="true" focusable="false" data-prefix="far" 
               data-icon="eye" class="svg-inline--fa fa-eye fa-icon fa-2x" 
               role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512">
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

                    {
                      // $(".btn_view_doc").attr("oncontextmenu","return false;");

                      // var table = $('#tbl_documents').DataTable({
                      //   responsive: true,
                      // });

                      // $('#tbl_documents tbody').on('click', '.btn_view_doc', async function () {
                      //   var data = table.row($(this).parents('tr')).data();
                      //   window.open(`${urlFile}`, '_blank');
                      // });



                    }

                    //sey nuvo

                    {

                    }


                    const btn_view_doc = document.getElementById(element.Id + '_view_doc');
                    const btn_view_doc_details = document.getElementById(element.Id + '_view_doc_details');



                    await btn_view_doc?.addEventListener('click', async (event) => {


                      $(".modal").css("display", "block");
                      window.open(`${urlFile}#toolbar=0`, '_blank');
                      // window.addEventListener("contextmenu", function (e) {
                      //   e.preventDefault();
                      // }, false);

                    });



                    //view details_doc
                    await btn_view_doc_details?.addEventListener('click', async () => {


                      //    window.open(`https://ncaircalin.sharepoint.com/sites/TestMyGed/SitePages/Document.aspx?document=${element.Title}&documentId=${element.FolderID}`, '_blank');


                      //   window.open(`${this.context.pageContext.web.absoluteUrl}/SitePages/Document.aspx?document=${element.Title}&documentId=${element.FolderID}`, '_blank');


                      window.open(`https://frcidevtest.sharepoint.com/sites/myGed/SitePages/Document.aspx?document=${element.Title}&documentId=${element.FolderID}`, '_blank');
                      //   window.location.replace(`https://frcidevtest.sharepoint.com/sites/myGed/SitePages/Home.aspx?folder=${element.FolderID}`);

                      //    window.history.pushState("data", "Title", `https://frcidevtest.sharepoint.com/sites/myGed/SitePages/Home.aspx?folder=${element.FolderID}`);

                      var urlFile_download = '';
                      var titleFolder = '';
                      var pdfNameDownload = '';

                      //getbyIDDocuments
                      const itemDoc: any = await sp.web.lists.getByTitle("Documents").items.getById(element.Id)();


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

                      console.log(itemDoc);


                      const allItemsFolder: any[] = await sp.web.lists.getByTitle('Documents').items.select("FolderID,Title").filter("FolderID eq '" + itemDoc.ParentID + "'").getAll();

                      allItemsFolder.forEach((x) => {

                        titleFolder = x.Title;

                      });


                      $("#input_type_doc").val(itemDoc.ParentID + "_" + titleFolder);
                      //  $("#input_type_doc").val(itemDoc.ParentID);
                      //input_type_doc
                      $("#input_number").val(itemDoc.Title);
                      $("#input_revision").val(itemDoc.revision);
                      $("#input_status").val(itemDoc.status);
                      $("#input_owner").val(itemDoc.owner);
                      $("#input_activeDate").val(itemDoc.active_date);
                      $("#input_filename").val(itemDoc.filename);
                      $("#input_author").val(itemDoc.author);
                      // $("#input_reviewDate").val(item1.);
                      $("#input_keywords").val(itemDoc.keywords);
                      $("#input_description").val(itemDoc.description);
                      $("#created_by").val(itemDoc.owner);

                      $("#updated_by").val(itemDoc.updateBy);
                      $("#updated_time").val(itemDoc.updatedDate);

                      //   $("#creation_date").val(itemDoc.Created);
                      $("#creation_date").val(itemDoc.createdDate);



                      console.log("BTN VIEW DETAIL ID", element.Id);

                      $("#access_form").css("display", "block");
                      $("#doc_form").css("display", "none");
                      $("#doc_details").css("display", "block");
                      $("#table_documents").css("display", "block");
                      $("#h2_title").text(element.Title);
                      $("#nav_file").css("display", "block");


                      //   $("#input_type_doc").val(element.parentID + "_" + titleFolder);
                      // $("#input_type_doc").val(element.ParentID);
                      // //input_type_doc
                      // $("#input_number").val(element.Title);
                      // $("#input_revision").val(element.revision);
                      // $("#input_status").val(element.status);
                      // $("#input_owner").val(element.owner);
                      // $("#input_activeDate").val(element.active_date);
                      // $("#input_filename").val(element.filename);
                      // $("#input_author").val(element.author);
                      // // $("#input_reviewDate").val(item1.);
                      // $("#input_keywords").val(element.keywords);
                      // $("#input_description").val(element.description);

                      //delete document
                      {
                        var delete_doc: Element = document.getElementById("delete_document");
                        let nav_html_delete_doc: string = '';
                        nav_html_delete_doc = `
                    <a href="#" title="Supprimer" 
                    role="button" id='${element.Id}_deleteDoc'>
                <svg aria-hidden="true" focusable="false" data-prefix="far" data-icon="trash-can" 
                class="svg-inline--fa fa-trash-can fa-icon fa-2x" role="img" xmlns="http://www.w3.org/2000/svg" 
                viewBox="0 0 448 512">
                <path fill="currentColor" d="M160 400C160 408.8 152.8 416 144 416C135.2 416 128 408.8 128 400V192C128 183.2 135.2 176 144 176C152.8 176 160 183.2 160 192V400zM240 400C240 408.8 232.8 416 224 416C215.2 416 208 408.8 208 400V192C208 183.2 215.2 176 224 176C232.8 176 240 183.2 240 192V400zM320 400C320 408.8 312.8 416 304 416C295.2 416 288 408.8 288 400V192C288 183.2 295.2 176 304 176C312.8 176 320 183.2 320 192V400zM317.5 24.94L354.2 80H424C437.3 80 448 90.75 448 104C448 117.3 437.3 128 424 128H416V432C416 476.2 380.2 512 336 512H112C67.82 512 32 476.2 32 432V128H24C10.75 128 0 117.3 0 104C0 90.75 10.75 80 24 80H93.82L130.5 24.94C140.9 9.357 158.4 0 177.1 0H270.9C289.6 0 307.1 9.358 317.5 24.94H317.5zM151.5 80H296.5L277.5 51.56C276 49.34 273.5 48 270.9 48H177.1C174.5 48 171.1 49.34 170.5 51.56L151.5 80zM80 432C80 449.7 94.33 464 112 464H336C353.7 464 368 449.7 368 432V128H80V432z"></path></svg> 
                    </a>
                    
                    `;

                        delete_doc.innerHTML = nav_html_delete_doc;

                        const btn_delete_doc = document.getElementById(element.Id + '_deleteDoc');


                        await btn_delete_doc?.addEventListener('click', async () => {
                          if (confirm(`Êtes-vous sûr de vouloir supprimer ${element.Title} ?`)) {

                            try {
                              var res = await sp.web.lists.getByTitle('Documents').items.getById(parseInt(element.Id)).delete()
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


                      }

                      //azoute permission
                      {
                        //add permission user


                        var add_user_permission_container: Element = document.getElementById("add_btn_user");

                        let add_btn_user_permission: string = `
      <button type="button" class="btn btn-primary add_group mb-2" id=${itemDoc.Id}_add_user>Ajouter</button>
      `;

                        add_user_permission_container.innerHTML = add_btn_user_permission;

                        const btn_add_user = document.getElementById(itemDoc.Id + '_add_user');

                        var peopleID = null;


                        await btn_add_user?.addEventListener('click', async () => {



                          const user: any = await sp.web.siteUsers.getByEmail($("#users_name").val().toString())();

                          const allItemsFolder: any[] = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID,Title,revision,IsFolder,description").filter("Title eq '" + itemDoc.Title + "'").getAll();


                          var response_same_doc = allItemsFolder;

                          console.log(response_same_doc);

                          users_Permission = user;

                          console.log("USERS FOR PERMISSION", users_Permission);


                          try {
                            console.log("KEY", item.key);

                            // response_same_doc.forEach(async (x) => {

                            await sp.web.lists.getByTitle("AccessRights").items.add({
                              Title: itemDoc.Title.toString(),
                              groupName: $("#users_name").val(),
                              permission: $("#permissions_user option:selected").val(),
                              FolderIDId: itemDoc.FolderID.toString(),
                              PrincipleID: user.Id
                            })
                              .then(() => {
                                // alert("Autorisation ajoutée à ce document avec succès.");
                              })
                              .then(() => {
                                // window.location.reload();
                              });

                            // });

                            // alert("Autorisation ajoutée à ce document avec succès.");
                            // window.location.reload();
                          }

                          catch (e) {
                            alert("Erreur: " + e.message);
                          }



                        });

                      }

                      //permission table 
                      //load table permission

                      {
                        var response = null;
                        let html: string = ``;

                        var permission_container: Element = document.getElementById("tbl_permission");
                        permission_container.innerHTML = "";


                        const allPermissions: any[] = await sp.web.lists.getByTitle('AccessRights').items.select("ID,groupName,permission,FolderIDId").filter("FolderIDId eq '" + itemDoc.Id + "'").getAll();

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
                               <button id="btn${element.ID}_edit" class='buttoncss' type="button"></button>
                               
                               
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

                      //update doc
                      {

                        var update_doc_container: Element = document.getElementById("btn_update_doc");

                        let update_btn_doc: string = `<button type="button" class="btn btn-primary update_details_doc" id='${itemDoc.Id}_update_details_doc'>Edit Details</button>`;

                        update_doc_container.innerHTML = update_btn_doc;


                        const btn_edit_doc = document.getElementById(itemDoc.Id + '_update_details_doc');

                        await btn_edit_doc?.addEventListener('click', async () => {

                          let user_current = await sp.web.currentUser();

                          let text = $("#input_type_doc").val();
                          const myArray = text.toString().split("_");
                          let parentId = myArray[0];


                          if (confirm(`Etes-vous sûr de vouloir mettre à jour les détails de ${itemDoc.Title} ?`)) {

                            try {

                              const i = await await sp.web.lists.getByTitle('Documents').items.add({
                                // const i = await await sp.web.lists.getByTitle('Documents').items.getById(parseInt(itemDoc.Id)).update({
                                Title: $("#input_number").val(),
                                description: $("#input_description").val(),
                                keywords: $("#input_keywords").val(),
                                doc_number: $("#input_number").val(),
                                revision: $("#input_revision").val(),
                                ParentID: parseInt(parentId),
                                FolderID: itemDoc.FolderID,
                                fileName: $("input_filename").val(),
                                IsFolder: "FALSE",
                                owner: itemDoc.owner,
                                updatedBy: user_current.Title,
                                createdDate: $("#creation_date").val(),
                                updatedDate: new Date().toLocaleString()
                              })
                                .then(async (iar) => {

                                  item = iar.data.ID;

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


                        });

                      }

                      //view doc

                      {
                        var view_doc: Element = document.getElementById("view_doc");


                        let nav_html_view_doc: string = '';


                        // console.log("ONSELECT", item.label);

                        nav_html_view_doc = `
                      <a href="#" role="button" title="Voir le document" id="${itemDoc.Id}_view_doc">
                      <svg aria-hidden="true" focusable="false" data-prefix="far" 
                      data-icon="eye" class="svg-inline--fa fa-eye fa-icon fa-2x" 
                      role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512">
                      <path fill="currentColor" d="M160 256C160 185.3 217.3 128 288 128C358.7 128 416 185.3 416 256C416 326.7 358.7 384 288 384C217.3 384 160 326.7 160 256zM288 336C332.2 336 368 300.2 368 256C368 211.8 332.2 176 288 176C287.3 176 286.7 176 285.1 176C287.3 181.1 288 186.5 288 192C288 227.3 259.3 256 224 256C218.5 256 213.1 255.3 208 253.1C208 254.7 208 255.3 208 255.1C208 300.2 243.8 336 288 336L288 336zM95.42 112.6C142.5 68.84 207.2 32 288 32C368.8 32 433.5 68.84 480.6 112.6C527.4 156 558.7 207.1 573.5 243.7C576.8 251.6 576.8 260.4 573.5 268.3C558.7 304 527.4 355.1 480.6 399.4C433.5 443.2 368.8 480 288 480C207.2 480 142.5 443.2 95.42 399.4C48.62 355.1 17.34 304 2.461 268.3C-.8205 260.4-.8205 251.6 2.461 243.7C17.34 207.1 48.62 156 95.42 112.6V112.6zM288 80C222.8 80 169.2 109.6 128.1 147.7C89.6 183.5 63.02 225.1 49.44 256C63.02 286 89.6 328.5 128.1 364.3C169.2 402.4 222.8 432 288 432C353.2 432 406.8 402.4 447.9 364.3C486.4 328.5 512.1 286 526.6 256C512.1 225.1 486.4 183.5 447.9 147.7C406.8 109.6 353.2 80 288 80V80z">
                      </path></svg>
                      
                      </a>`;

                        view_doc.innerHTML = nav_html_view_doc;

                        const btn_view_doc = document.getElementById(itemDoc.Id + '_view_doc');

                        await btn_view_doc?.addEventListener('click', async () => {
                          window.open(`${urlFile_download}`, '_blank');
                        });

                      }

                      //download doc

                      // {
                      //   var download_doc: Element = document.getElementById("download_doc");


                      //   let nav_html_download_doc: string = '';


                      //   // console.log("ONSELECT", item.label);

                      //   nav_html_download_doc = `
                      //                   <a href="#" title="Telecharger le document" 
                      //                   role="button" id='${itemDoc.Id}_download_doc' >
                      //                   <svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="download" 
                      //                   class="svg-inline--fa fa-download fa-icon fa-2x" role="img" xmlns="http://www.w3.org/2000/svg"
                      //                    viewBox="0 0 512 512"><path fill="currentColor" d="M288 32c0-17.7-14.3-32-32-32s-32 14.3-32 32V274.7l-73.4-73.4c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3l128 128c12.5 12.5 32.8 12.5 45.3 0l128-128c12.5-12.5 12.5-32.8 0-45.3s-32.8-12.5-45.3 0L288 274.7V32zM64 352c-35.3 0-64 28.7-64 64v32c0 35.3 28.7 64 64 64H448c35.3 0 64-28.7 64-64V416c0-35.3-28.7-64-64-64H346.5l-45.3 45.3c-25 25-65.5 25-90.5 0L165.5 352H64zM432 456c-13.3 0-24-10.7-24-24s10.7-24 24-24s24 10.7 24 24s-10.7 24-24 24z">
                      //                    </path>
                      //                    </svg>
                      //                   </a>

                      //                   `;

                      //   download_doc.innerHTML = nav_html_download_doc;

                      //   const btn_download_doc = document.getElementById(itemDoc.Id + '_download_doc');

                      //   await btn_download_doc?.addEventListener('click', async () => {

                      //     try {

                      //       {


                      //         const user = await sp.web.currentUser();



                      //         const dateDownload = Date();

                      //         const textWatermark = 'Uncontrolled Copy - Downloaded on ' + dateDownload + ' .';
                      //         //  Uncontrolled Copy - Downloaded on" the you put the date

                      //         //Load PDF
                      //         const existingPdfBytes = await fetch(urlFile_download).then(res => res.arrayBuffer());
                      //         const pdfDoc = await PDFDocument.load(existingPdfBytes);
                      //         console.log('pdfDoc Starting...');

                      //         const pages = await pdfDoc.getPages();

                      //         for (const [i, page] of Object.entries(pages)) {
                      //           const firstPage = pages[0];

                      //           const { width, height } = firstPage.getSize();

                      //           const helveticaFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
                      //           const fontSize = 16;

                      //           page.drawText(textWatermark, {
                      //             x: 60,
                      //             y: 60,
                      //             size: fontSize,
                      //             font: helveticaFont,
                      //             color: rgb(1, 0, 1),
                      //             opacity: 0.4,
                      //             rotate: degrees(55)

                      //           });

                      //         }

                      //         const pdfBytes = await pdfDoc.save();

                      //         console.log('pdfBytes: ', pdfBytes);

                      //         download(pdfBytes, [pdfNameDownload], "application/pdf");
                      //       }

                      //     }
                      //     catch (e) {

                      //       alert("Cannot download this file for the following reason: " + e.message);
                      //     }



                      //   });

                      // }


                      //so bne lezot versions
                      {
                        //display so table
                        // $("#table_version_doc").css("display", "block");

                        var html_document_versions: string = ``;
                        var response_doc_versions = null;
                        var value1 = "FALSE";

                        var document_versions_container: Element = document.getElementById("tbl_documents_versions_bdy");
                        document_versions_container.innerHTML = "";

                        const all_documents_versions: any[] = await sp.web.lists.getByTitle('Documents').items
                          .select("Id,ParentID,FolderID,Title,revision,IsFolder,description")
                          .filter("Title eq '" + itemDoc.Title + "' and IsFolder eq '" + value1 + "'")
                          .get();

                        response_doc_versions = all_documents_versions;

                        if (response_doc_versions.length > 0) {
                          $("#table_version_doc").css("display", "block");
                          await response_doc_versions.forEach(async (element_version) => {


                            html_document_versions += `
                            <tr>
                            <td class="text-left">${element_version.Id}</td>
            
                            <td class="text-left">${element_version.Title}</td>
            
                            <td class="text-left"> 
                            ${element_version.revision}          
                            </td>
            
                            
                            <td>

                           <a href="#"  title="Voir le document" id="${element_version.Id}_view_doc_version" class="btn_view_doc" style="padding-left: inherit;">
                           <svg aria-hidden="true" focusable="false" data-prefix="far" 
                           data-icon="eye" class="svg-inline--fa fa-eye fa-icon fa-2x" 
                           role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512">
                           <path fill="currentColor" d="M160 256C160 185.3 217.3 128 288 128C358.7 128 416 185.3 416 256C416 326.7 358.7 384 288 384C217.3 384 160 326.7 160 256zM288 336C332.2 336 368 300.2 368 256C368 211.8 332.2 176 288 176C287.3 176 286.7 176 285.1 176C287.3 181.1 288 186.5 288 192C288 227.3 259.3 256 224 256C218.5 256 213.1 255.3 208 253.1C208 254.7 208 255.3 208 255.1C208 300.2 243.8 336 288 336L288 336zM95.42 112.6C142.5 68.84 207.2 32 288 32C368.8 32 433.5 68.84 480.6 112.6C527.4 156 558.7 207.1 573.5 243.7C576.8 251.6 576.8 260.4 573.5 268.3C558.7 304 527.4 355.1 480.6 399.4C433.5 443.2 368.8 480 288 480C207.2 480 142.5 443.2 95.42 399.4C48.62 355.1 17.34 304 2.461 268.3C-.8205 260.4-.8205 251.6 2.461 243.7C17.34 207.1 48.62 156 95.42 112.6V112.6zM288 80C222.8 80 169.2 109.6 128.1 147.7C89.6 183.5 63.02 225.1 49.44 256C63.02 286 89.6 328.5 128.1 364.3C169.2 402.4 222.8 432 288 432C353.2 432 406.8 402.4 447.9 364.3C486.4 328.5 512.1 286 526.6 256C512.1 225.1 486.4 183.5 447.9 147.7C406.8 109.6 353.2 80 288 80V80z">
                           </path></svg>
            
                           </a>
            
                            </td>
                          
                           `;



                          });


                        }

                        document_versions_container.innerHTML += html_document_versions;




                      }




                    });

                    $("#edit_cancel_doc").click(() => {
                      $("#table_documents").css("display", "block");
                      $("#doc_form").css("display", "none");
                      $("#access_form").css("display", "block");
                      $("#edit_details").css("display", "none");
                      $("#access_rights_form").css("display", "none");
                      $("#notifications_doc_form").css("display", "none");
                      $(".dossier_headers").css("display", "block");


                    });

                    $("#edit_cancel").click(() => {

                      $("#edit_details").css("display", "none");

                    });

                    //change href value
                    // const a_view_doc = document.getElementById(element.Id + '_view_doc');

                    // $(a_view_doc).attr("href", urlFile);



                  });

                console.log("URL FILE", urlFile);


                // <!-- <button id="btn${element.ID}_remove" class='buttoncss' type="button">Voir</button> -->
                // html_document += `
                // <td class="text-left">${urlFile}</td>
                // </tr>`;


              });


              document_container.innerHTML += html_document;


              // var table = $('#tbl_documents').DataTable({
              //   responsive: true,
              //   columnDefs: [{
              //     visible: false,
              //     targets: [0],
              //     searchable: false
              //   }]
              // });

              // $('#tbl_documents tbody').on('click', '.btn_view_doc', async function () {

              //   var data = table.row($(this).parents('tr')).data();
              //   window.open(data[3], '_blank');



              // });



              // $("#tbl_documents").DataTable();


              // $('#tbl_documents').on('click', '.btn_view_doc', (event) => {
              //   var data = $("#tbl_documents").DataTable({ retrieve: true }).row($(event.currentTarget).parents('tr')).data();
              //   window.open(data[3], '_blank');
              // });

            }

            else {
              $("#alert_0_doc").css("display", "block");
              $("#table_documents").css("display", "none");
            }

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

            //    const all_folders: any = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID,Title,IsFolder,description").top(5000).filter("ParentID eq '" + item.key + "'").get();


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

              //  users_Permission.forEach((result: ISiteUserInfo) => {

              //   peopleID = result.Id;

              // });

              // if ($("#permissions_user option:selected").val() == "NONE") {

              //   var docResponse = null;
              //   // try {
              //   console.log("RESPONSE ROLE DEF ID", roleDefID);


              //   // }
              //   // catch (e) {

              //   //   alert("Exception" + e.message);
              //   // }

              // }

              // else {


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

        onSelect={async () => {

          this.onSelect(this.state.TreeLinks);
        }}



      >

        {
          < FontAwesomeIcon icon={item.icon} className="fa-icon" ></FontAwesomeIcon >
        }

        &nbsp;

        {item.label}

      </span>
    );


  }

  private async loadDocsFromFolders(id: any) {

    //render table
    {

      var response_doc = null;
      var response_distinc = [];
      var html_document: string = ``;
      var value1 = "FALSE";

      var pdfName = '';


      var document_container: Element = document.getElementById("tbl_documents_bdy");
      document_container.innerHTML = "";


      const all_documents: any[] = await sp.web.lists.getByTitle('Documents').items
        .select("Id,ParentID,FolderID,Title,revision,IsFolder,description")
        .filter("ParentID eq '" + x + "' and IsFolder eq '" + value1 + "'")
        .get();

      response_doc = all_documents;


      var result = response_doc.filter((obj, pos, arr) => {
        return arr.map(mapObj =>
          mapObj.Title).lastIndexOf(obj.Title) == pos;
      });

      console.log("ALL", response_doc);

      console.log("RESULT DISTINCT", result);
      console.log("RESULT DISTINCT ARRAY LOT LA", response_distinc);


      if (result.length > 0) {


        html_document = ``;
        $("#alert_0_doc").css("display", "none");
        $("#table_documents").css("display", "block");



        await result.forEach(async (element) => {



          var urlFile = '';
          html_document += `
                    <tr>
                    <td class="text-left">${element.Id}</td>
    
                    <td class="text-left">${element.Title}</td>
    
                    <td class="text-left"> 
                    ${element.description}          
                    </td>
    
                    
                    <td>
                    <a href="#" title="Mettre à jour le document" role="button" id="${element.Id}_view_doc_details" class="btn_view_doc_details">
                    <svg aria-hidden="true" focusable="false" data-prefix="far" 
                    data-icon="pen-to-square" class="svg-inline--fa fa-pen-to-square fa-icon fa-2x" 
                    role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">
                    <path fill="currentColor" d="M373.1 24.97C401.2-3.147 446.8-3.147 474.9 24.97L487 37.09C515.1 65.21 515.1 110.8 487 138.9L289.8 336.2C281.1 344.8 270.4 351.1 258.6 354.5L158.6 383.1C150.2 385.5 141.2 383.1 135 376.1C128.9 370.8 126.5 361.8 128.9 353.4L157.5 253.4C160.9 241.6 167.2 230.9 175.8 222.2L373.1 24.97zM440.1 58.91C431.6 49.54 416.4 49.54 407 58.91L377.9 88L424 134.1L453.1 104.1C462.5 95.6 462.5 80.4 453.1 71.03L440.1 58.91zM203.7 266.6L186.9 325.1L245.4 308.3C249.4 307.2 252.9 305.1 255.8 302.2L390.1 168L344 121.9L209.8 256.2C206.9 259.1 204.8 262.6 203.7 266.6zM200 64C213.3 64 224 74.75 224 88C224 101.3 213.3 112 200 112H88C65.91 112 48 129.9 48 152V424C48 446.1 65.91 464 88 464H360C382.1 464 400 446.1 400 424V312C400 298.7 410.7 288 424 288C437.3 288 448 298.7 448 312V424C448 472.6 408.6 512 360 512H88C39.4 512 0 472.6 0 424V152C0 103.4 39.4 64 88 64H200z"></path></svg></a>
    
    
    
                   <a href="#"  title="Voir le document" id="${element.Id}_view_doc"  class="btn_view_doc" style="padding-left: inherit;">
                   <svg aria-hidden="true" focusable="false" data-prefix="far" 
                   data-icon="eye" class="svg-inline--fa fa-eye fa-icon fa-2x" 
                   role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512">
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
                window.open(`${urlFile}#toolbar=0`, '_blank');

              });



              //view details_doc
              await btn_view_doc_details?.addEventListener('click', async () => {


                window.open(`https://frcidevtest.sharepoint.com/sites/myGed/SitePages/Document.aspx?document=${element.Title}&documentId=${element.Id}`, '_blank');

              });


            });

          console.log("URL FILE", urlFile);




        });


        document_container.innerHTML += html_document;
      }

      else {
        $("#alert_0_doc").css("display", "block");
        $("#table_documents").css("display", "none");
      }

    }


  }


}


