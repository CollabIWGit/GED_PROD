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


import 'bootstrap/dist/css/bootstrap.css';


// import Form from 'react-bootstrap/Form';
// import Button from 'react-bootstrap/Button';

require('./../../../common/css/common.css');
require('./../../../common/css/sidebar.css');
require('./../../../common/css/pagecontent.css');



export default class MyGedTreeView extends React.Component<IMyGedTreeViewProps, IMyGedTreeViewState> {



  constructor(props: IMyGedTreeViewProps) {



    super(props);

    sp.setup({
      spfxContext: this.props.context
      //props.context
    });

    // const sp = spfi().using(SPFx(this.props.context));
    this.state = {
      TreeLinks: [],

    };




    this._getLinks(sp);

    this.getParentID(this.getItemId());

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
    var tree: ITreeItem[] = [];

    // var treearr: any = [];
    // var treeSub: any = [];
    // var tree: any = [];



    allItems.forEach((v, i) => {

      //ParentId ==> ParentID


      if (v["ParentID"] == -1) {

        var str = v["Title"];

        const tree: ITreeItem = {
          // v.id ==> v.FolderID
          id: v["ID"],
          key: v["FolderID"],
          label: str,
          data: 0,
          icon: faFolder,
          children: [],
          revision: ""

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
          revision: ""

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

      else if (v["ParentID"] !== -1 && v["IsFolder"] === "FALSE") {

        // console.log("We have a file here with ParentId :" + v["ParentID"]);


        //v["ServerRedirectedEmbedUrl"] ==> v["FileUrl"]
        const iconName: IconProp = faFile;

        const tree: ITreeItem = {
          // v.id ==> v.FolderID
          id: v["ID"],
          key: v["FolderID"],
          label: v["Title"] + "-" + v["revision"],
          icon: faFile,
          data: v["FileUrl"],
          revision: v['revision']
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

    var remainingArr = treearr.filter(data => data.data == 0);



    //  var x = remainingArr.filter((a, i) => remainingArr.findIndex((s) => a.label === s.age) === i);

    // var remainingArr = treearr.filter(data => data.data == 0).map(item => item.label).filter((value, index, self) => self.indexOf(value) === index);

    // const key = 'label';
    // const arrayUniqueByKey = [...new Map(remainingArr.map(item =>
    //   [item[key], item])).values()];

    tree = remainingArr;





    // console.log("DISTINCT ITEMS", remainingArr);

    this.setState({ TreeLinks: remainingArr });
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



    await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + id + "'").get().then((results) => {
      parentID = results[0].ParentID;
      parentIDArray.push(parentID);

      console.log("Parent 1", parentID);

    });


    while (parentID != null) {
      await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + parentID + "'").get().then((results) => {
        parentID = results[0].ParentID;
        parentIDArray.push(parentID);

        console.log("Parent 2", parentID);
      });
    }


    //parentIDArray.push(parseInt(this.getItemId()));
    parentIDArray.sort(function (a, b) { return a - b });
    console.log("ArrayParent", parentIDArray);

    //return parentIDArray;
  }



  public render(): React.ReactElement<IMyGedTreeViewProps> {

    parentIDArray.unshift(1, parseInt(this.getItemId()));


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
                    defaultExpandedChildren={false}

                    defaultExpandedKeys={parentIDArray.sort(function (a, b) { return a - b })}
                    // selectionMode={TreeViewSelectionMode.None}
                    selectChildrenIfParentSelected={false}
                    showCheckboxes={true}
                    treeItemActionsDisplayMode={TreeItemActionsDisplayMode.Buttons}
                    onSelect={this.onSelect}
                    onExpandCollapse={this.onTreeItemExpandCollapse}
                    onRenderItem={this.renderCustomTreeItem}
                    defaultSelectedKeys={[parseInt(this.getItemId())]}
                    expandToSelected={true}
                  />
                </div>
              </div>
            </div>

          </div>



          <div className="col-sm-9">

            <form id="form_metadata">

              <h2 id='h2_title'></h2>

              <div className="form-row">

                <div className="form-group col-md-6">
                  <Label>Title
                    <input type="text" className="form-control" placeholder="First name" />
                  </Label>
                </div>

                <div className="form-group col-md-6">
                  <Label>Title
                    <input type="email" className="form-control" id="input_title" />
                  </Label>
                </div>

              </div>

              <div className="row">
                <div className="col-6">
                  <input type="text" className="form-control" placeholder="First name" />
                </div>
                <div className="col-6">
                  <input type="text" className="form-control" placeholder="Last name" />
                </div>
              </div>

              //coumans la

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

              //2ieme

              <div className="row">
                <div className="col-3">
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
                  <div className="col-3">
                  <Label>Document Number
                    <input type="text" id='input_number' className='form-control' />
                  </Label>
                  </div>
              </div>


              <div className='form-group col-6'>

                <Label>Title
                  <input type="email" className="form-control" id="input_title" />
                </Label>

              </div>

              <div className='form-group col-3'>

                <Label>Type
                  <input type="email" className="form-control" id="input_type" />
                </Label>

              </div>

              <div className='form-group col-3'>

                <Label>Document Number
                  <input type="text" id='input_number' className='form-control' />
                </Label>

              </div>

              <div className='form-group col-3'>

                <Label>
                  Revision
                  <input type="text" id='input_revision' className='form-control' />
                </Label>

              </div>

              <div className='form-group col-3'>

                <Label>
                  Status
                  <input type="text" id='input_status' className='form-control' />
                </Label>

              </div>

              <div className='form-group col-3'>

                <Label>
                  Owner
                  <input type="text" id='input_owner' className='form-control' />
                </Label>

              </div>

              <div className='form-group col-3'>

                <Label>
                  Active Date
                  <input type="text" id='input_activeDate' className='form-control' />
                </Label>

              </div>

              <div className='form-group col-6'>

                <Label>
                  Filename
                  <input type="text" id='input_filename' className='form-control' />
                </Label>

              </div>

              <div className='form-group col-3'>

                <Label>
                  Author
                  <input type="text" id='input_author' className='form-control' />
                </Label>

              </div>

              <div className='form-group col-3'>

                <Label>
                  Review Date
                  <input type="text" id='input_reviewDate' className='form-control' />
                </Label>

              </div>

              <div className='form-group col-6'>

                <Label>
                  Keywords
                  <textarea id='input_keywords' className='form-control' rows={3} />
                </Label>

              </div>

              <button type="button" className="btn btn-primary mb-2" id='view'>View Document</button>





            </form>


            {/* 
            <form id="form_metadata">
              <h2 className="h2_title"></h2>

              <label>
                Title
                <input type="text" id='input_title' className='form-control' />
              </label>
              <label>
                Type
                <input type="text" id='input_type' className='form-control' />
              </label>
              <label>
                Document Number
                <input type="text" id='input_number' className='form-control' />
              </label>
              <label>
                Revision
                <input type="text" id='input_revision' className='form-control' />
              </label>
              <label>
                Status
                <input type="text" id='input_status' className='form-control' />
              </label>

              <label>
                Owner
                <input type="text" id='input_owner' className='form-control' />
              </label>

              <label>
                Active Date
                <input type="text" id='input_activeDate' className='form-control' />
              </label>

              <label>
                Filename
                <input type="text" id='input_filename' className='form-control' />
              </label>

              <label>
                Author
                <input type="text" id='input_author' className='form-control' />
              </label>

              <label>
                Review Date
                <input type="text" id='input_reviewDate' className='form-control' />
              </label>

              <label>
                Keywords
                <input type="text" id='input_keywords' className='form-control' />
              </label>

              
            </form> */}


          </div>
        </div>
      </div>


    );

  }

  // expandNode(key: any) {
  //   this.treeView.expandItem(key);
  // }
  // collapseNode(key: any) {
  //   this.treeView.collapseItem(key);
  // }



  private async onTreeItemSelect(items: ITreeItem[]) {

    items.forEach((item: any) => {
      console.log("Items selected: ", item.label);
    })

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
          console.log("DATA value", item.data);
          if (item.data == 1 || item.data == 0) {
          }

          else {

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
