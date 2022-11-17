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

import { sp, List, IItemAddResult, UserCustomActionScope, Items, Item } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { getIconClassName, Label } from 'office-ui-fabric-react';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faFolder, faFolderOpen, faFileWord } from '@fortawesome/free-regular-svg-icons'
import { faFile } from '@fortawesome/free-solid-svg-icons'
import { IconName, IconProp } from '@fortawesome/fontawesome-svg-core';
import { useEffect } from 'react';


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
    });

    // const sp = spfi().using(SPFx(this.props.context));
    this.state = {
      TreeLinks: []
    };

    this._getLinks(sp);

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

        };

        treearr.push(tree);
        console.log("Tree 1", tree);

      }

      // v["FileSystemObjectType"] ==> v["IsFolder"]
      else if (v["ParentID"] !== -1 && v["IsFolder"] === "TRUE") {

        console.log("We have a sub folder here.");
        var str = v["Title"];

        const tree: ITreeItem = {
          // v.id ==> v.FolderID
          id: v["ID"],
          key: v["FolderID"],
          // label: str.normalize('NFD').replace(/\p{Diacritic}/gu, ""),
          label: str,
          data: 1,
          icon: faFolderOpen,
          children: [],

        };

        console.log("Tree 2", tree);


        // ParentID ==> FolderID
        var treecol: Array<ITreeItem> = treearr.filter((value) => { return value.key == v["ParentID"]; });


        if (treecol.length != 0) {

          treearr.push(tree);
          console.log("TREE COL", treecol);
          console.log("COL SUB", treeSub);
          treecol[0].children.push(tree);
        }
      }

      else if (v["ParentID"] !== -1 && v["IsFolder"] === "FALSE") {

        console.log("We have a file here with ParentId :" + v["ParentID"]);


        //v["ServerRedirectedEmbedUrl"] ==> v["FileUrl"]
        const iconName: IconProp = faFile;

        const tree: ITreeItem = {
          // v.id ==> v.FolderID
          id: v["ID"],
          key: v["FolderID"],
          label: v["Title"],
          icon: faFile,
          data: v["FileUrl"],

        };


        // ParentID ==> FolderID
        var treecol: Array<ITreeItem> = treearr.filter((value) => { return value.key == v["ParentID"]; });
        if (treecol.length != 0) {
          treecol[0].children.push(tree);
        }
      }

    });
    console.log("TREE ARRAY", treearr);
    console.log("TREE ARRAY SUB", treeSub);

    var remainingArr = treearr.filter(data => data.data == 0);

    tree = remainingArr;

    this.setState({ TreeLinks: remainingArr });
  }



  public render(): React.ReactElement<IMyGedTreeViewProps> {

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
                    defaultExpandedKeys={['1']}
                    // selectionMode={TreeViewSelectionMode.None}
                    selectChildrenIfParentSelected={false}
                    showCheckboxes={true}
                    treeItemActionsDisplayMode={TreeItemActionsDisplayMode.Buttons}
                    onSelect={this.onSelect}
                    onExpandCollapse={this.onTreeItemExpandCollapse}
                    onRenderItem={this.renderCustomTreeItem}
                    defaultExpandedChildren={false}

                  />


                </div>
              </div>
            </div>

          </div>



          <div className="col-sm-9">

            <form id="form_metadata">
              <label>
                Title
                <input type="text" id='input_title' className='form-control' />
              </label>
              <label>
                Url
                <input type="text" id='input_url' className='form-control' />
              </label>
              <label>
                Url
                <input type="text" id='input_url' className='form-control' />
              </label>
              <label>
                Url
                <input type="text" id='input_url' className='form-control' />
              </label>
              <label>
                Url
                <input type="text" id='input_url' className='form-control' />
              </label>

              <button type="button" id='Interne' onClick={() => {
                console.log("Button Clicked!!");


              }

              }>Expand Internal</button>
            </form>


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
      console.log("Items selected: ", item.label);
    })
  }


  private onTreeItemExpandCollapse(item: ITreeItem, isExpanded: boolean) {
    console.log((isExpanded ? "Item expanded: " : "Item collapsed: ") + item.label);
    console.log(item.key);
    console.log(item.key);

    $("#text").text(item.label);
  }


  private renderCustomTreeItem(item: ITreeItem): JSX.Element {

    return (
      <span
        //onclick
        onClick={async event => {
          console.log("DATA value", item.data);
          if (item.data == 1 || item.data == 0) {
          }

          else {

            //   const item_detail: any = await sp.web.lists.getByTitle("TestDocument").items.getById(parseInt(item.id))();
            const item_detail: any = await sp.web.lists.getByTitle("Documents").items.getById(parseInt(item.id))();

            Object.keys(item_detail).forEach((key) => {
              $("#input_title").val(item_detail.Title);
              $("#input_url").val(item_detail.FileUrl);
            });

            console.log("DETAAAAILS", item_detail);
          }
          // console.log(this.getItemMetadata(item.id));
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
