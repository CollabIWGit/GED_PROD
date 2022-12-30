import * as React from 'react';

import styles from './MyGedTreeView.module.scss';
import { IMyGedTreeViewProps, IMyGedTreeViewState } from './IMyGedTreeView';
import { escape } from '@microsoft/sp-lodash-subset';
import { TreeView, ITreeItem, TreeViewSelectionMode, TreeItemActionsDisplayMode, ITreeItemActions } from "@pnp/spfx-controls-react/lib/TreeView";
// import { SPFI, spfi, SPFx } from "@pnp/sp";
// import { Container, Row, Col, Card, Form, Nav } from "react-bootstrap";
import 'bootstrap/dist/css/bootstrap.min.css';
import $ from 'jquery';
import Popper from 'popper.js';
import 'bootstrap/dist/js/bootstrap.bundle.min';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { sp, List, IItemAddResult, UserCustomActionScope, Items, Item } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { getIconClassName, Label } from 'office-ui-fabric-react';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faFolder, faFolderOpen, faFileWord } from '@fortawesome/free-regular-svg-icons'
import { faFile } from '@fortawesome/free-solid-svg-icons'
import { IconName, IconProp, parse } from '@fortawesome/fontawesome-svg-core';
import { useEffect, useState } from 'react';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { IAttachmentInfo } from "@pnp/sp/attachments";
import "@pnp/sp/attachments";
import { IItem } from "@pnp/sp/items/types";



var parentIDArray = [];

var sorted = [];

var val = [];



import 'bootstrap/dist/css/bootstrap.css';
import 'bootstrap/dist/css/bootstrap.min.css';



// import Form from 'react-bootstrap/Form';
// import Button from 'react-bootstrap/Button';

require('./../../../common/css/common.css');
require('./../../../common/css/sidebar.css');
require('./../../../common/css/pagecontent.css');
require('./../../../common/css/spinner.css');




export default class MyGedTreeView extends React.Component<IMyGedTreeViewProps, IMyGedTreeViewState> {



  constructor(props: IMyGedTreeViewProps) {



    super(props);

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


    // this._getLinks(sp);
    this._getLinks2(sp);

    this.render();


    // this.getFirstParent();


    //var node = this.getParent();

    console.log("NODES", parentIDArray);

  }

  private async _getLinks(sp) {

    // const allItems: any[] = await sp.web.lists.getByTitle("TestTreeView").items();
    // const allItems: any[] = await sp.web.lists.getByTitle("TestTreeView1").items.getAll();
    //const allItems: any[] = await sp.web.lists.getByTitle("TestDocument").items.getAll();
    const allItems: any[] = await sp.web.lists.getByTitle("Documents").items.getAll();



    console.log("ALL ITEMS", allItems);


    var treearr: ITreeItem[] = [];
    var treeSub: ITreeItem[] = [];
    var tree: any = [];

    // var treearr: any = [];
    // var treeSub: any = [];
    // var tree: any = [];





    allItems.forEach((v, i) => {

      //ParentId ==> ParentID


      if (v["ParentID"] == -1) {


        console.log("We have the main folder here.");

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
          file: "No"

        };

        treearr.push(tree);
        // console.log("Tree 1", tree);

      }

      // v["FileSystemObjectType"] ==> v["IsFolder"]
      else if (v["ParentID"] !== -1 && v["IsFolder"] === "TRUE") {


        console.log("We have a sub folder here.");
        var str = v["Title"];

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
          file: "No"
        };

        // console.log("Tree 2", tree);


        // ParentID ==> FolderID
        var treecol: Array<ITreeItem> = treearr.filter((value) => { return value.key == v["ParentID"]; });


        if (treecol.length != 0) {

          treearr.push(tree);
          // console.log("TREE COL", treecol);
          // console.log("COL SUB", treeSub);
          treecol[0].children.push(tree);
        }
      }

      // else if (v["ParentID"] !== -1 && v["IsFolder"] === "FALSE") {
      else if (v["IsFolder"] === "FALSE") {


        // console.log("We have a file here with ParentId :" + v["ParentID"]);

        console.log("We have a file here.");


        //v["ServerRedirectedEmbedUrl"] ==> v["FileUrl"]
        const iconName: IconProp = faFile;

        const tree: ITreeItem = {
          // v.id ==> v.FolderID
          id: v["ID"],
          key: v["FolderID"],
          label: v["Title"] + "-" + v["revision"],
          icon: faFile,
          data: v["FileUrl"],
          revision: v['revision'],
          file: "Yes"

        };


        // ParentID ==> FolderID
        var treecol: Array<ITreeItem> = treearr.filter((value) => { return value.key == v["ParentID"]; });


        if (treecol.length != 0) {
          treecol[0].children.push(tree);
        }
      }
    });



    // console.log("TREE ARRAY", treearr);
    // console.log("TREE ARRAY SUB", treeSub);

    //var remainingArr = treearr.filter(data => (data.data == 0 || data.data == 1));

    var remainingArr = treearr.filter(data => data.data == 0);




    //  var x = remainingArr.filter((a, i) => remainingArr.findIndex((s) => a.label === s.age) === i);

    // var remainingArr = treearr.filter(data => data.data == 0).map(item => item.label).filter((value, index, self) => self.indexOf(value) === index);

    // const key = 'label';
    // const arrayUniqueByKey = [...new Map(remainingArr.map(item =>
    //   [item[key], item])).values()];

    tree = remainingArr;

    console.log("LENGTH", remainingArr.length);


    // console.log("DISTINCT ITEMS", remainingArr);

    this.setState({ TreeLinks: remainingArr });

    // this.setState({ TreeLinks: treearr });

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


    const allItemsMain: any[] = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID,Title,IsFolder").filter("IsFolder eq '" + value1 + "'").getAll();
    const allItemsFile: any[] = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID,Title,revision,IsFolder").filter("IsFolder eq '" + value2 + "'").getAll();

    const allItemsMain_sorted: any[] = allItemsMain.sort((a, b) => { return a.Title - b.Title });
    // const allItemsMain_sorted: any[] = allItemsMain.sort();


    console.log("ARRRANGED ARRAY: " + allItemsMain_sorted);

    // const allItemsMain_sorted: any[] = allItemsMain.sort((a, b) => a.FolderID - b.FolderID);
    var x = 0;



    allItemsMain_sorted.forEach(v => {


      if (v["ParentID"] == -1) {

        console.log("We have the main folder here.");

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
          file: "No"

        };



        treearr.push(tree);
        // console.log("Tree 1", tree);

      }

      // v["FileSystemObjectType"] ==> v["IsFolder"]
      else {

        allKeys.push(v["FolderID"]);

        console.log("We have a sub folder here.");
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
          children: [],
          revision: "",
          file: "No"
        };

        // console.log("Tree 2", tree);

       

        // ParentID ==> FolderID



        //  var treecol: Array<any> = treearr.filter((value) => { return value.key === v["ParentID"]; }).sort((a, b) => {return a.label - b.label} );

        var treecol: Array<any> = treearr.filter((value) => { return value.key === v["ParentID"]; });



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

      allItemsMain_sorted.forEach(x => {

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
            file: "No"

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


    console.log("KEYS MISSING LENGTH", keysMissing.length);
    console.log("KEYS MISSING", keysMissing);
    console.log("KEYS PRESENT LENGTH", keysPresent.length);
    console.log("KEYS MISSING LENGTH", 763 - keysPresent.length);
    console.log("COUNTER", counter);
    console.log("COUNTERSUB", counterSUB);
    console.log("TREE ARRAY LENGTH", treearr.length);

    allItemsFile.forEach((v) => {
      console.log("We have a file here.");


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
        file: "Yes"

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

  private getItemId() {
    var queryParms = new URLSearchParams(document.location.search.substring(1));
    var myParm = queryParms.get("folder");
    if (myParm) {
      return myParm.trim();
    }
  }

  // private async getParentID(id: any) {

  //   var parentID = null;
  //   var counter = 1;



  //   try {
  //     await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + id + "'").get().then((results) => {

  //       parentID = results[0].ParentID;
  //       //  parentIDArray.unshift(parseInt(parentID));
  //       parentIDArray.push(parseInt(parentID));

  //       console.log("Parent 1", parentID);
  //     });


  //     //  while (parentID != null || parentID != undefined) {
  //     while (parentID != -1) {


  //       await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + parentID + "'").get().then((results) => {
  //         parentID = results[0].ParentID;
  //         //  parentIDArray.unshift(parseInt(parentID));
  //         parentIDArray.push(parseInt(parentID));

  //       });

  //       counter = counter + 1;
  //       console.log(`Parent ${counter}`, parentID);

  //     }

  //     parentIDArray.push(parseInt(this.getItemId()));


  //     //parentIDArray.push(parseInt(this.getItemId()));
  //     // parentIDArray.sort(function (a, b) { return a - b });
  //     console.log("ArrayParent", parentIDArray);


  //     parentIDArray.forEach(x => {

  //       val.push(x);

  //     });



  //     //return parentIDArray;

  //   }
  //   catch (e) {
  //     console.log(e.message);
  //   }



  // }

  private async getParentID(id: any) {

    var parentID = null;



    await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + id + "'").get().then((results) => {
      parentID = results[0].ParentID;
      parentIDArray.push(parentID);

      console.log("Parent 1", parentID);

    });


    while (parentID != -1) {
      await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + parentID + "'").get().then((results) => {
        parentID = results[0].ParentID;
        parentIDArray.unshift(parentID);

        console.log("Parent 2", parentID);
      });
    }


    parentIDArray.push(parseInt(this.getItemId()));
    // parentIDArray.sort(function (a, b) { return a - b });
    console.log("ArrayParent", parentIDArray);

    //return parentIDArray;
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
                    // defaultExpandedKeys={parentIDArray}
                    defaultExpandedKeys={parentIDArray}

                    defaultExpandedChildren={false}


                    selectChildrenIfParentSelected={false}
                    showCheckboxes={true}
                    treeItemActionsDisplayMode={TreeItemActionsDisplayMode.Buttons}
                    onSelect={this.onSelect}
                    onExpandCollapse={this.onTreeItemExpandCollapse}
                    onRenderItem={this.renderCustomTreeItem}
                    // defaultSelectedKeys={[parseInt(this.getItemId())]}
                    defaultSelectedKeys={[parseInt(x)]}
                    expandToSelected={true}
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
                <h2 id='h2_title'></h2>

                <div className="row">
                  <div className="col-6">
                    <Label>Title
                      <input type="email" className="form-control" id="input_title" />
                    </Label>
                  </div>
                  <div className="col-3">
                    <Label>Type
                      <input type="email" className="form-control" id="input_type" />
                    </Label>
                  </div>
                  <div className="col-3">
                    <Label>Document Number
                      <input type="text" id='input_number' className='form-control' />
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
                  <div className="col-3">
                    <Label>
                      Author
                      <input type="text" id='input_author' className='form-control' />
                    </Label>
                  </div>
                  <div className="col-3">

                    <button type="button" className="btn btn-primary mb-2" id='view' >View Document</button>

                  </div>
                </div>


                <div className="row">
                  <div className="col-8">
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

              </div>

              <div id="access_form">

                <h2 id='h2_folderName'></h2>

                <nav aria-label="breadcrumb" id='nav'>
                  <ol className="breadcrumb">
                    <li className="breadcrumb-item"><a href="#" role="button" onClick={(event: React.MouseEvent<HTMLElement>) => { $("#access_rights_form").css("display", "block"); $("#subfolders_form").css("display", "none"); }}>Access Rights</a></li>
                    <li className="breadcrumb-item"><a href="#" role="button" onClick={(event: React.MouseEvent<HTMLElement>) => { $("#access_rights_form").css("display", "none"); $("#subfolders_form").css("display", "block"); }}>Add Subfolders</a></li>
                  </ol>
                </nav>

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
                    <div className="col-3">
                      <button type="button" className="btn btn-primary mb-2" id='add_group'>Add group</button>
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

                    <div className="col-3">
                      <button type="button" className="btn btn-primary mb-2" id='add_subfolder'>Add subfolder</button>
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

  private async addSubfolders(item: ITreeItem) {



    console.log("ID", item.id);

  }


  private async onTreeItemSelect(items: ITreeItem[]) {

    items.forEach((item: any) => {
      console.log("Items selected: ", item.label);
    });

  }

  private onSelect(items: ITreeItem[]) {
    items.forEach((item: ITreeItem) => {

      item.iconProps.color = "black";

    })
  }


  private onTreeItemExpandCollapse(item: ITreeItem, isExpanded: boolean) {
    console.log((isExpanded ? "Item expanded: " : "Item collapsed: ") + item.label);
    console.log(item.key);
    console.log(item.key);

    $("#text").text(item.label);
  }


  private renderCustomTreeItem(item: ITreeItem): JSX.Element {


    var docResponse = null;

    return (
      <span
        //onclick
        onClick={async event => {
          console.log("DATA value", item.label);

          if (item.data == 1 || item.data == 0) {
            $("#h2_folderName").text(item.label);


            $("#access_form").css("display", "block");

            $("#doc_form").css("display", "none");

            $("#add_subfolder").click(async function () {
              console.log("ID", item.id);

              await sp.web.lists.getByTitle("Documents").items.add({
                Title: $("#folder_name").val(),
                ParentID: item.key,
                IsFolder: "TRUE"
              })
                .then(async (iar) => {

                  const list = sp.web.lists.getByTitle("Documents");

                  await list.items.getById(iar.data.ID).update({
                    FolderID: parseInt(iar.data.ID)
                  });
                });

            });

            $("#add_group").click(async function () {

              await sp.web.lists.getByTitle("AccessRights").items.add({
                Title: item.label,
                groupName: $("#group_name").val(),
                permission: $("#permissions option:selected").val(),
                FolderIDId: item.key
              }).then(() =>{
                console.log("Permission added to this folder.")
              })

            });

          }


          else {

            $("#access_form").css("display", "none");

            $("#doc_form").css("display", "block");

            var urlFile = '';

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
              })


            //  var item1: any = sp.web.lists.getByTitle('Documents').items.getById(parseInt(item.id));
            const item1: any = await sp.web.lists.getByTitle("Documents").items.getById(parseInt(item.id))();
            console.log(item1);

            console.log(item1);

            Object.keys(item1).forEach((key) => {

              $("#input_title").val(item1.Title);
              $("#input_type").val(item1.type);
              $("#input_number").val(item1.doc_number);
              $("#input_revision").val(item1.revision);
              $("#input_status").val(item1.status);
              $("#input_owner").val(item1.owner);
              $("#input_activeDate").val(item1.active_date);
              $("#input_filename").val(item1.filename);
              $("#input_author").val(item1.author);
              // $("#input_reviewDate").val(item1.);
              $("#input_keywords").val(item1.keywords);
              $("#h2_title").text(item1.Title);

              document.getElementById('view').onclick = function () {
                window.open(`${urlFile}`, '_blank');
              };

            });



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


