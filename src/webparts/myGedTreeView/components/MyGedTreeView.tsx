
import { MSGraphClient } from '@microsoft/sp-http';
import { IMyGedTreeViewProps, IMyGedTreeViewState } from './IMyGedTreeView';
import { TreeView, ITreeItem, TreeViewSelectionMode, TreeItemActionsDisplayMode } from "@pnp/spfx-controls-react/lib/TreeView";
// import 'bootstrap/dist/css/bootstrap.min.css';
import $, { event } from 'jquery';
import Popper from 'popper.js';
// import 'bootstrap/dist/js/bootstrap.bundle.min';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { sp, List, IItemAddResult, UserCustomActionScope, Items, ITerm, ISiteGroup, ISiteGroupInfo, SPRest, RoleAssignment, Item, RoleDefinition } from "@pnp/sp/presets/all";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { getIconClassName, Label, rgb2hex } from 'office-ui-fabric-react';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faFolder, faFolderOpen, faFileWord, faEdit, faTrashCan, faBell, faEye, faStar, faSquareCheck, faBookmark } from '@fortawesome/free-regular-svg-icons'
import { faFile, faLock, faFolderPlus, faDownload, faMagnifyingGlass, faDeleteLeft, faCircleInfo, faSquareXmark, faSquareCaretLeft, faCircleXmark } from '@fortawesome/free-solid-svg-icons'
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

import Form from 'react-bootstrap/Form';
import { degrees, PDFDocument, radians, rgb, rotateDegrees, rotateRadians, StandardFonts, } from 'pdf-lib/cjs/api';
import * as download from 'downloadjs';
import "@pnp/sp/security/web";
import "@pnp/sp/site-users/web";

SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.25/css/jquery.dataTables.min.css');
SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.8.335/pdf.js'),
  SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.8.335/pdf.worker.js')


import { Group } from '@microsoft/microsoft-graph-types';
import pdfjsLib from "pdfjs-dist";
// import * as pdfjsLib from "pdfjs-dist";



declare namespace pdfjsViewer {
  class PDFViewer {
    constructor(options: any);
    setDocument(url: string);
    currentScaleValue: string;
  }
}


declare global {
  interface Window {
    pdfjsLib: any;
  }
}



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
var department;

// var remainingArr: any = [];
var myVar;
var x;

interface MyObject {
  PrincipleId: string;
  role: string;
  [key: string]: any;
}


// import 'bootstrap/dist/css/bootstrap.css';
// import 'bootstrap/dist/css/bootstrap.min.css';
import { ITreeViewState } from '@pnp/spfx-controls-react/lib/controls/treeView/ITreeViewState';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import { max } from 'lodash';
import { Client } from '@microsoft/microsoft-graph-client';


// import Form from 'react-bootstrap/Form';
// import Button from 'react-bootstrap/Button';

// const js = fs.readFileSync('path/to/your/script.js', 'utf8');

// require('./../../../common/css/common.css');
// require('./../../../common/css/sidebar.css');
// require('./../../../common/css/pagecontent.css');
// require('./../../../common/css/spinner.css');
// require('./../../../common/css/responsive.css');
// require('./../../../common/js/jquery.min');

import 'datatables.net';
import * as moment from 'moment';
import 'downloadjs';


var department;

export default class MyGedTreeView extends React.Component<IMyGedTreeViewProps, IMyGedTreeViewState, any> {

  private graphClient: MSGraphClient;
  readonly context: WebPartContext;

  constructor(props: IMyGedTreeViewProps, context: WebPartContext) {

    super(props, context);

    sp.setup({
      spfxContext: this.props.context
    });


    this.props.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        this.graphClient = client;
      });

    this.state = {
      selectedKey: null,
      TreeLinks: [],
      parentIDArray: [],
      isLoaded: false,
      isToggledOn: false,
      isToggleOnDept: false
    };


    this.onSelect = this.onSelect.bind(this);
    this.handleIconClick = this.handleIconClick.bind(this);
    this.handleIconClickDept = this.handleIconClickDept.bind(this);
    this.toggleIcon = this.toggleIcon.bind(this);
    // Bind the toggleIcon function to the current component instance
  }


  async handleIconClick() {
    this.setState(prevState => ({
      isToggledOn: !prevState.isToggledOn
    }));

    var x = this.getDossierID();
    var y = this.getDossierTitle();

    var url = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${x}`;


    try {

      if (!this.state.isToggledOn) {
        await this.addBookmark(x, y);
        alert("Vous avez ajouté ce document comme favori.");
        window.location.href = url;

      }
      else {
        await this.removeBookmark(x);
        alert("Vous avez supprimé ce document des favoris.");
        window.location.href = url;

      }
    } catch (error) {
      alert("Failed to update bookmark: " + error);
    }


  }

  async handleIconClickDept() {
    this.setState(prevState => ({
      isToggleOnDept: !prevState.isToggleOnDept
    }));

    var x = this.getDossierID();
    var y = this.getDossierTitle();

    var url = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${x}`;
    // var url = `${this.context.pageContext.web.absoluteUrl}/SitePages/documentation.aspx?folder=${x}`;

    try {

      if (!this.state.isToggleOnDept) {
        await this.addDept(x, y);
        alert("You have entered this folder in department list.");
        window.location.href = url;
      }
      else {
        await this.removeDept(x);
        alert("You have removed this folder in department list.");
        window.location.href = url;
      }
    } catch (error) {
      alert("Failed to update list: " + error);
    }
  }


  onSelect(item) {
    this.setState({ selectedKey: item.key });
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

  private async _getLinks3(sp) {
    // Retrieve all items from the "Documents" list with the "IsFolder" field set to "TRUE"
    const allItems: any[] = await sp.web.lists.getByTitle('Documents').items
      .select("ID,ParentID,FolderID,Title,IsFolder,description")
      .filter("IsFolder eq 'TRUE'")
      .getAll();

    // Create a flat map of all items
    const itemsMap = new Map<number, ITreeItem>();
    allItems.forEach(item => {
      const treeItem: ITreeItem = {
        id: item.ID,
        key: item.FolderID,
        label: item.Title,
        data: 0,
        icon: faFolder,
        children: [],
        revision: "",
        file: "No",
        description: item.description,
        parentID: item.ParentID
      };
      itemsMap.set(treeItem.key, treeItem);
    });

    // Build the tree structure
    const rootItems: any[] = [];


    itemsMap.forEach(item => {
      if (item.parentID === -1) {
        rootItems.push(item);
      } else {
        const parentItem = itemsMap.get(item.parentID);
        if (parentItem) {
          parentItem.children.push(item);
       //   parentItem.children.sort((a, b) => a.label.substr(0, 3).localeCompare(b.label.substr(0, 3))); // Sort children alphabetically by label
          parentItem.children.sort((a, b) => a.label.localeCompare(b.label)); // Sort children alphabetically by label
        } else {
          rootItems.push(item); // Add item to root if parent not found
        }
      }
    });

    // Sort the tree structure alphabetically with 3 characters
    // const sortedTreeArr = rootItems.map((tree) => {
    //   if (tree.children) {
    //     tree.children.sort((a, b) => a.label.substr(0, 3).localeCompare(b.label.substr(0, 3)));
    //   }
    //   return tree;
    // }).sort((a, b) => a.label.localeCompare(b.label));

    const sortedTreeArr = rootItems.map((tree) => {
      if (tree.children) {
        tree.children.sort((a, b) => a.label.localeCompare(b.label));
      }
      return tree;
    }).sort((a, b) => a.label.localeCompare(b.label));
    



    return sortedTreeArr;
  }

  private getItemId() {
    var queryParms = new URLSearchParams(document.location.search.substring(1));
    var myParm = queryParms.get("folder");
    if (myParm) {
      return myParm.trim();
    }
  }

  // private async getAllGroups(graphClient: MSGraphClient): Promise<any[]> {
  //   try {
  //     const groups = await graphClient.api('/groups')
  //       .version('v1.0')
  //       .get();

  //     return groups.value;
  //   } catch (error) {
  //     console.error(error);
  //     return [];
  //   }
  // }

  private async getAllGroups(graphClient: MSGraphClient): Promise<Group[]> {
    const allGroups: Group[] = [];

    let nextPageUrl = '/groups';
    while (nextPageUrl) {
      const response = await graphClient.api(nextPageUrl).version('v1.0').get();
      allGroups.push(...response.value);
      nextPageUrl = response["@odata.nextLink"] ?? null;
    }

    return allGroups;
  }

  private async getAllGroups2(graphClient: MSGraphClient): Promise<Group[]> {
    const groupIds = ["20", "17", "12"]; // Specify the group IDs to retrieve
    const allGroups: Group[] = [];

    let nextPageUrl = '/groups';
    while (nextPageUrl) {
      const response = await graphClient.api(nextPageUrl).version('v1.0').get();
      const filteredGroups = response.value.filter((group: Group) => groupIds.includes(group.description));
      allGroups.push(...filteredGroups);
      nextPageUrl = response["@odata.nextLink"] ?? null;
    }

    return allGroups;
  }

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

  public async generateTable(groups: any, x) {
    {

      var value2 = 'TRUE';

      const folderInfo = await sp.web.lists.getByTitle('Documents').items
        .select("ID,ParentID,FolderID,Title,revision,IsFolder,description,attachmentUrl,IsFiligrane,IsDownloadable, inheriting")
        .top(5000)
        .filter(`FolderID eq '${x}' and IsFolder eq '${value2}'`)
        .getAll();

      var permission_container: Element = document.getElementById("spListPermissions");

      // while (permission_container.firstChild) {
      //   permission_container.removeChild(permission_container.firstChild);
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

        //  if (element1.role !== "NONE") {
        html += `
          <tr>
          <td class="text-left" id="${element1.id}">${element1.id}</td>

          <td class="text-left" id="${element1.id}_personName">${element1.title}</td>
          
          <td class="text-center" id="${element1.id}_permission_value"> ${element1.role} </td>

          <td class="text-center">
          <a id="btn${element1.id}_edit" class='buttoncss' role="button">
          
          <svg aria-hidden="true" focusable="false" data-prefix="far" data-icon="trash-can" class="svg-inline--fa fa-trash-can fa-icon fa-2x" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512">
                    <path fill="currentColor" d="M160 400C160 408.8 152.8 416 144 416C135.2 416 128 408.8 128 400V192C128 183.2 135.2 176 144 176C152.8 176 160 183.2 160 192V400zM240 400C240 408.8 232.8 416 224 416C215.2 416 208 408.8 208 400V192C208 183.2 215.2 176 224 176C232.8 176 240 183.2 240 192V400zM320 400C320 408.8 312.8 416 304 416C295.2 416 288 408.8 288 400V192C288 183.2 295.2 176 304 176C312.8 176 320 183.2 320 192V400zM317.5 24.94L354.2 80H424C437.3 80 448 90.75 448 104C448 117.3 437.3 128 424 128H416V432C416 476.2 380.2 512 336 512H112C67.82 512 32 476.2 32 432V128H24C10.75 128 0 117.3 0 104C0 90.75 10.75 80 24 80H93.82L130.5 24.94C140.9 9.357 158.4 0 177.1 0H270.9C289.6 0 307.1 9.358 317.5 24.94H317.5zM151.5 80H296.5L277.5 51.56C276 49.34 273.5 48 270.9 48H177.1C174.5 48 171.1 49.34 170.5 51.56L151.5 80zM80 432C80 449.7 94.33 464 112 464H336C353.7 464 368 449.7 368 432V128H80V432z">
                    </path></svg>
          
          </a>
        </td>
          

          </tr>
          `;

        //  }

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

        // var table = $('#tbl_permission').DataTable({
        //   order: [0, 'desc'],
        //   columnDefs: [{
        //     targets: [8],
        //     orderable: false,
        //   },
        //   {
        //     targets: [0],
        //     visible: false,
        //   }]
        // });

        $('#tbl_permission tbody').on('click', '.buttoncss', async (event) => {
          var data = table.row($(event.currentTarget).parents('tr')).data();
          alert("Remove permission with people id" + data[0]);

          try {

            console.log("KEY", folderInfo[0].FolderID);

            var x = await this.getChildrenById(folderInfo[0].FolderID, []);


            await sp.web.lists.getByTitle("AccessRights").items.add({
              Title: folderInfo[0].Title.toString(),
              groupName: $("#users_name").val(),
              permission: "NONE",
              FolderID: folderInfo[0].ID.toString(),
              PrincipleID: data[0]
              //  RoleDefID: permission
            })
              .then(async () => {
                await Promise.all(x.map(async (item_group) => {
                  await sp.web.lists.getByTitle("AccessRights").items.add({
                    Title: item_group.Title.toString(),
                    groupName: $("#users_name").val(),
                    permission: "NONE",
                    FolderID: item_group.ID,
                    PrincipleID: data[0]
                  });
                }));

              })
              .then(async () => {
                alert("Autorisation supprimée avec succès.");
                await sp.web.lists.getByTitle("Documents").items.getById(folderInfo[0].ID).update({
                  inheriting: "NO"
                }).then(result => {
                  console.log("Item updated successfully");
                }).catch(error => {
                  console.log("Error updating item: ", error);
                });

                window.location.reload();

              });
          }
          catch (e) {
            console.log(e.message);
          }

          // await this.downloadDoc(data[2], data[5], data[6], 'ARCHIVED COPY - Downloaded on ');
        });

      } else {
        // Container is not empty
      }


      // await Promise.all(filteredPermissions.map(async (element1) => {
      //   const deleteButton = document.getElementById(element1.Id + '_edit');

      //   deleteButton?.addEventListener('click', async () => {

      //     const user: any = await sp.web.siteUsers.getByEmail(element1.groupName)();


      //     try {
      //       console.log("KEY", item.key);

      //       await sp.web.lists.getByTitle("AccessRights").items.add({
      //         Title: item.label.toString(),
      //         groupName: element1.groupName,
      //         //   groupName: "zpeerbaccus.ext@aircalin.nc",
      //         permission: "NONE",
      //         FolderID: item.key.toString(),
      //         PrincipleID: user.Id,


      //         // PrincipleID: 15

      //       })
      //         .then(() => {
      //           alert("Autorisation supprimer à ce dossier avec succès.");
      //         })
      //         .then(() => {
      //           window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${item.key}`;
      //         });
      //     }

      //     catch (e) {
      //       alert("Erreur: " + e.message);
      //     }

      //   });


      // }));

    }
  }

  public async getListItemPermissions(siteUrl, listName, itemId, username, password) {
    try {
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

  // loadPdfJs() {
  //   return new Promise<void>((resolve, reject) => {
  //     const script = document.createElement("script");
  //     script.type = "text/javascript";
  //     script.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.9.359/pdf.min.js";
  //     script.onload = () => {
  //       window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.9.359/pdf.worker.min.js";
  //       resolve();
  //     };
  //     script.onerror = reject;
  //     document.head.appendChild(script);
  //   });
  // }

  async componentDidMount() {

    // var graphClient: MSGraphClient;

    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        this.graphClient = client;
      });

    var x = this.getItemId();

    // this.loadPdfJs();



    if (x == null || x == undefined || x == "") {
      x = "1";
    }

    const loader = document.createElement("div");
    loader.id = "loader_else";
    loader.style.display = "none";
    loader.style.position = "fixed";
    loader.style.top = '0';
    loader.style.left = '0';
    loader.style.width = "100%";
    loader.style.height = "100%";
    loader.style.backgroundColor = "white";
    loader.style.zIndex = "9999";
    loader.style.display = "flex";
    loader.style.justifyContent = "center";
    loader.style.alignItems = "center";

    const image = document.createElement("img");
    image.src = "https://ncaircalin.sharepoint.com/:i:/r/sites/MyGed/SiteAssets/images/logoGed.png?csf=1&web=1&e=CTrOpq";
    image.alt = "Loading...";
    image.style.transform = "scale(1)";
    image.style.animation = "pulsate 1.5s ease-in-out infinite";

    loader.appendChild(image);

    document.body.style.overflow = "hidden";
    document.body.appendChild(loader);
    loader.style.display = "flex";

    const style = document.createElement("style");
    style.innerHTML = `
    @keyframes pulsate {
      0% {
        transform: scale(1);
      }
      50% {
        transform: scale(1.2);
      }
      100% {
        transform: scale(1);
      }
    }
    `;
    document.head.appendChild(style);


    var value2 = 'TRUE';

    const folderInfo = await sp.web.lists.getByTitle('Documents').items
      .select("ID,ParentID,FolderID,Title,revision,IsFolder,description,attachmentUrl,IsFiligrane,IsDownloadable, inheriting")
      .top(5000)
      .filter(`FolderID eq '${x}' and IsFolder eq '${value2}'`)
      .getAll();

    console.log("LOADED");

    await this.getBasePermTest2('df095fdf-9978-44f8-941c-23b6c095751a', folderInfo[0].ID)
      .then(async result => {
        // Handle the result
        console.log('High Value:', result.high);
        console.log('Low Value:', result.low);

        const high = result.high;
        const low = result.low;

        if ((high == 2147483647 && low == 4294967295)  || (high == 2147483647 && low == 4294705151)) { //full control
          console.log("You have full control!");
          const { permissions } = await this.getListItemPermissions('https://ncaircalin.sharepoint.com/sites/MyGed', "Documents", folderInfo[0].ID, "mgolapkhan.ext@aircalin.nc", "musharaf2897");
          await this.generateTable(permissions, Number(x));
       
        }
        else if (high == 432 && low == 1011030767) { //edit
          $("#ajouterDept, #accesFolder, #bouton_delete, #editFolder, #addFolder").css("display", "none");
        }
        else if (high == 176 && low == 138612833) { //read
          $("#nav").css("display", "none");
        }

        else {

        }
      })
      .catch(error => {
        // Handle any errors
        console.error('Error:', error);
      });


    // try {
    //   const isAdmin = await this.checkIfUserIsAdmin(this.graphClient);
    //   const isRefUser = isAdmin || await checkIfUserIsRefUser(this.graphClient);
    //   const isGuestUser = isAdmin || await checkIfUserIsGuestUser(this.graphClient);

    //   if (isAdmin) {
    //     console.log('User is an administrator.');
    //     // const { permissions, groupPermissions } = await getListItemPermissions('https://ncaircalin.sharepoint.com/sites/MyGed', "Documents", item.id, "mgolapkhan.ext@aircalin.nc", "musharaf2897");

    //     const { permissions } = await this.getListItemPermissions('https://ncaircalin.sharepoint.com/sites/MyGed', "Documents", folderInfo[0].ID, "mgolapkhan.ext@aircalin.nc", "musharaf2897");

    //     await this.generateTable(permissions, Number(x));
    //     console.log("PERMISSIONS ON ITEM", permissions);

    //   } else if (isRefUser) {
    //     console.log('User is a MYGED_REF user.');
    //     $("#ajouterDept, #accesFolder, #bouton_delete, #editFolder, #addFolder").css("display", "none");


    //   } else if (isGuestUser) {
    //     console.log('User is a MYGED_GUEST user.');
    //     $("#nav").css("display", "none");


    //   } else {

    //     console.log('User is not an administrator or a MYGED_REF or MYGED_GUEST user.');
    //     $("#nav").css("display", "block");

    //     const { permissions } = await this.getListItemPermissions('https://ncaircalin.sharepoint.com/sites/MyGed', "Documents", folderInfo[0].ID, "mgolapkhan.ext@aircalin.nc", "musharaf2897");
    //     await this.generateTable(permissions, Number(x));
    //     console.log("PERMISSIONS ON ITEM", permissions);

    //   }
    // } catch (error) {
    //   console.log('An error occurred while checking user permissions:', error);
    // }

    async function checkIfUserIsRefUser(graphClient: MSGraphClient): Promise<boolean> {
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

    async function checkIfUserIsGuestUser(graphClient: MSGraphClient): Promise<boolean> {
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

    if (x == null || x == undefined || x == "") {
      // const allItems = await this._getLinks2(sp);
      // const allItems = await this._getFolders(sp);




      const allItems = await this._getLinks3(sp);
      this.setState({ TreeLinks: allItems });

      await this.fetchDocuments(1,);


      // console.log("COUNT MAIN", allItems);

      var xxx = await this.getAllGroups2(this.graphClient);
      var yy = await this.getGroupsByName(this.graphClient, "myGed");

      console.log("ALL AD MATCHING ", xxx);
      console.log("ALL AD GROUPS", yy);

      loader.style.display = "none";

    }

    else {

      const parentIDs = await this.getParentID(x);
      const allItems = await this._getLinks3(sp);

      this.setState({ parentIDArray: parentIDs, TreeLinks: allItems });

      const user = await sp.web.currentUser();


      var items = await sp.web.lists.getByTitle("Marque_Pages").items
        .select("ID")
        .filter(`FolderID eq '${x}' and user eq '${user.Title}'`)
        .get();

      if (items.length === 0) {
        this.setState({ isToggledOn: false });
      } else {
        this.setState({ isToggledOn: true });
      }

      var itemsDept = await sp.web.lists.getByTitle("Department").items
        .select("ID")
        .filter(`FolderID eq '${x}'`)
        .get();

      if (itemsDept.length === 0) {
        this.setState({ isToggleOnDept: false });
      } else {
        this.setState({ isToggleOnDept: true });
      }

      await this.fetchDocuments(Number(x));

      loader.style.display = "none";

      // Hide the loader and blur elements

    }

    //




    this.setState({ isLoaded: true });



  }

  private async add_permission_group2(group_name: string, permission: any, id: any, foldertitle: any, folderid: any, inherit: any, principleIdOfGroup: any) {



    try {

      var x = await this.getChildrenById(id, []);

      await sp.web.lists.getByTitle("AccessRights").items.add({
        Title: foldertitle.toString(),
        groupName: group_name,
        permission: $("#permissions_group option:selected").val(),
        FolderID: folderid,
        PrincipleID: principleIdOfGroup,
        groupTitle: group_name,
        RoleDefID: permission
      })

        .then(async () => {

          await Promise.all(x.map(async (item_group) => {

            await sp.web.lists.getByTitle("AccessRights").items.add({
              Title: item_group.Title.toString(),
              groupName: group_name,
              permission: $("#permissions_group option:selected").val(),
              FolderID: item_group.ID,
              PrincipleID: principleIdOfGroup,
              groupTitle: group_name,
              RoleDefID: permission
            });

          }));

        });


      alert("Autorisation ajoutée avec succès..");
      window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${id}`;
    }
    catch (e) {
      alert("Error: " + e.message);
    }

  }

  private async getChildrenById(id, items) {

    const children = await sp.web.lists.getByTitle("Documents").items
      .select("ID, Title, ParentID, inheriting, FolderID")
      .filter(`ParentID eq '${id}'`)
      .get();

    let result = [];

    for (const child of children) {
      result.push(child);
      const subChildren = await this.getChildrenById(child.FolderID, items);

      result = [...result, ...subChildren];
    }

    return result;
  }

  private async fetchDocuments(itemKey: number): Promise<void> {

    console.log("LOADED FETCH");
    let response_doc: any = null;
    let response_distinc: any[] = [];
    let html_document = '';
    let value1 = "FALSE";
    let value2 = "TRUE";
    let value3 = "";
    let principleIdOfGroup = "";

    let pdfName = '';

    console.log("ITEM KEY", itemKey);

    const folderInfo = await sp.web.lists.getByTitle('Documents').items
      .select("ID,ParentID,FolderID,Title,revision,IsFolder,description,attachmentUrl,IsFiligrane,IsDownloadable, inheriting")
      .top(5000)
      .filter(`FolderID eq '${itemKey}' and IsFolder eq '${value2}'`)
      .getAll();


    {
      {
        $("#access_form").css("display", "block");
        $("#doc_form").css("display", "none");
        $(".dossier_headers").css("display", "block");

        $("#subfolders_form").css("display", "none");

        $("#access_rights_form").css("display", "none");
        $("#notifications_doc_form").css("display", "none");

        $("#doc_details_add").css("display", "none");
        $("#edit_details").css("display", "none");
        $("#h2_folderName").text(folderInfo[0].Title);
      }

      $("#h2_folderName").text(folderInfo[0].Title);
    }

    {
      //render metadata
      {
        var fileName = "";
        var content = null;

        var filename_add = "";
        var content_add = null;
        var istrue = "TRUE";

        if (folderInfo[0].ParentID !== null) {
          const allItemsFolder: any[] = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID,Title,revision,IsFolder,description").filter("FolderID eq '" + folderInfo[0].ParentID + "' and IsFolder eq '" + istrue + "'").getAll();
          $("#parent_folder").val(allItemsFolder[0].FolderID + "_" + allItemsFolder[0].Title);

        }
        else {
          $("#parent_folder").val("");

        }


        $("#folder_name1").val(folderInfo[0].Title);
        $("#folder_desc").val(folderInfo[0].description);
      }

      //check checkboxes

      {
        const checkbox_fili = document.querySelector('input[name="checkFiligrane"]') as HTMLInputElement;
        checkbox_fili.checked = true;

        const checkbox_Imprimab = document.querySelector('input[name="checkImprimab"]') as HTMLInputElement;
        checkbox_Imprimab.checked = true;
      }
      // enta encore
      //bouton delete dossier
      {
        var delete_dossier: Element = document.getElementById("bouton_delete");


        let nav_html_delete_dossier: string = '';


        // console.log("ONSELECT", item.label);

        // nav_html_delete_dossier = `
        //                         <a href="#" title="Supprimer" 
        //                         role="button" id='${folderInfo[0].ID}_deleteFolder' style="color: rgb(13, 110, 253);">
        //                     <svg aria-hidden="true" focusable="false" data-prefix="far" data-icon="trash-can" 
        //                     class="svg-inline--fa fa-trash-can fa-icon fa-2x" role="img" xmlns="http://www.w3.org/2000/svg" 
        //                     viewBox="0 0 448 512">
        //                     <path fill="currentColor" d="M160 400C160 408.8 152.8 416 144 416C135.2 416 128 408.8 128 400V192C128 183.2 135.2 176 144 176C152.8 176 160 183.2 160 192V400zM240 400C240 408.8 232.8 416 224 416C215.2 416 208 408.8 208 400V192C208 183.2 215.2 176 224 176C232.8 176 240 183.2 240 192V400zM320 400C320 408.8 312.8 416 304 416C295.2 416 288 408.8 288 400V192C288 183.2 295.2 176 304 176C312.8 176 320 183.2 320 192V400zM317.5 24.94L354.2 80H424C437.3 80 448 90.75 448 104C448 117.3 437.3 128 424 128H416V432C416 476.2 380.2 512 336 512H112C67.82 512 32 476.2 32 432V128H24C10.75 128 0 117.3 0 104C0 90.75 10.75 80 24 80H93.82L130.5 24.94C140.9 9.357 158.4 0 177.1 0H270.9C289.6 0 307.1 9.358 317.5 24.94H317.5zM151.5 80H296.5L277.5 51.56C276 49.34 273.5 48 270.9 48H177.1C174.5 48 171.1 49.34 170.5 51.56L151.5 80zM80 432C80 449.7 94.33 464 112 464H336C353.7 464 368 449.7 368 432V128H80V432z"></path></svg> 
        //                         </a>`;

        nav_html_delete_dossier = `
                                <a title="Archiver le dossier" 
                                role="button" id='${folderInfo[0].ID}_deleteFolder' style="color: rgb(13, 110, 253);">
                                <img src="https://icons.iconarchive.com/icons/fa-team/fontawesome/128/FontAwesome-Box-Archive-icon.png" width="34" height="34" style="margin-top: -16px;" id="archiver_dossier">
                                </img>
                                </a>`;



        delete_dossier.innerHTML = nav_html_delete_dossier;

        const btn = document.getElementById(folderInfo[0].ID + '_deleteFolder');

        await btn?.addEventListener('click', async () => {
          // this.domElement.querySelector('#btn' + item.Id + '_edit').addEventListener('click', () => {
          //localStorage.setItem("contractId", item.Id);
          if (confirm(`Voulez-vous vraiment archiver ce dossier : ${folderInfo[0].Title} ?`)) {

            try {

              const list = sp.web.lists.getByTitle("Documents");

              const i = await list.items.getById(folderInfo[0].ID).update({
                ParentID: 791,
              })
                .then(() => {
                  alert("Dossier archivé avec succès.");
                })
                .then(() => {
                  window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx`;
                  // window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].ParentID}`;
                });


              // var res = await sp.web.lists.getByTitle('Documents').items.getById(parseInt(folderInfo[0].ID)).delete()
              //   .then(() => {
              //     alert("Dossier supprimé avec succès.");
              //   })
              //   .then(() => {
              //     window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx`;

              //     // window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].ParentID}`;
              //   });
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

        let update_btn_dossier: string = `<button type="button" class="btn btn-primary btn_edit_dossier" id='${folderInfo[0].ID}_update_details' style="font-size: 1em;">Modifier</button>
                `;

        update_dossier_container.innerHTML = update_btn_dossier;


        const btn_edit_dossier = document.getElementById(folderInfo[0].ID + '_update_details');

        await btn_edit_dossier?.addEventListener('click', async () => {


          let text = $("#parent_folder").val();
          const myArray = text.toString().split("_");
          let parentId = myArray[0];

          if (confirm(`Etes-vous sûr de vouloir mettre à jour les détails de ${folderInfo[0].Title} ?`)) {

            try {

              const i = await await sp.web.lists.getByTitle('Documents').items.getById(parseInt(folderInfo[0].ID)).update({
                Title: $("#folder_name1").val(),
                description: $("#folder_desc").val(),
                ParentID: parseInt(parentId)

              })
                .then(() => {
                  alert("Détails mis à jour avec succès");
                })
                .then(() => {
                  window.open(`https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].FolderID}`, "blank");
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
                <button type="button" class="btn btn-primary add_doc" id="${folderInfo[0].ID}_add_doc" style="font-size: 1em;">Sauvegarder</button>
                `;

        add_doc_container.innerHTML = add_btn_document;


        const btn_add_doc = document.getElementById(folderInfo[0].ID + '_add_doc');

        await btn_add_doc?.addEventListener('click', async () => {

          if ($("#input_revision_add").val() == "") {
            alert("Veuillez mettre une révision avant de continuer.")
          }

          else {
            const getCheckboxValue = (checkbox: HTMLInputElement): string => {
              return checkbox.checked ? "YES" : "NO";
            }

            const checkbox_Fili = document.querySelector<HTMLInputElement>('input[name="checkFiligrane"]');
            const checkbox_Imprimab = document.querySelector<HTMLInputElement>('input[name="checkImprimab"]');

            const value_fili = getCheckboxValue(checkbox_Fili);
            const value_impri = getCheckboxValue(checkbox_Imprimab);



            // const checkbox = document.getElementById(checkboxId);
            // if (checkbox.checked) {
            //   return checkbox.value;
            // } else {
            //   return null;
            // }

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
                    ParentID: folderInfo[0].FolderID,
                    IsFolder: "FALSE",
                    keywords: $("#input_keywords_add").val(),
                    owner: user_current.Title,
                    createdDate: new Date().toLocaleString(),
                    IsFiligrane: value_fili,
                    IsDownloadable: value_impri
                  })
                    .then(async (iar) => {

                      //   item = iar.data.ID;

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

                        })

                        .then(async () => {
                          await sp.web.lists.getByTitle("InheritParentPermission").items.add({
                            Title: folderInfo[0].Title,
                            FolderID: iar.data.ID,
                            IsDone: "NO",
                            ParentID: Number(folderInfo[0].FolderID)
                          });

                        });

                    })
                    .then(() => {
                      alert("Document creer avec succès");
                      window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].FolderID}`;
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




        });

      }

      //bouton add subfolder
      {
        var add_subfolder_container: Element = document.getElementById("add_btn_subFolder");

        let add_btn_subfolder: string = `
                <button type="button" class="btn btn-primary add_subfolder mb-2" id="${folderInfo[0].ID}_add_btn_subfolder" style="float: right; font-size: 1em;">Ajouter un sous-dossier</button>
                `;

        add_subfolder_container.innerHTML = add_btn_subfolder;


        const btn_add_subfolder = document.getElementById(folderInfo[0].ID + '_add_btn_subfolder');


        await btn_add_subfolder?.addEventListener('click', async () => {
          var subId = null;

          if ($("#folder_name").val() == '') {
            alert("Veuillez mettre une révision avant de continuer.")
          }

          else {
            try {
              await sp.web.lists.getByTitle("Documents").items.add({
                Title: $("#folder_name").val(),
                ParentID: folderInfo[0].FolderID,
                IsFolder: "TRUE"
              })
                .then(async (iar) => {

                  const list = sp.web.lists.getByTitle("Documents");

                  subId = iar.data.ID;

                  await list.items.getById(iar.data.ID).update({
                    FolderID: parseInt(iar.data.ID),

                  })
                    .then(async () => {

                      await sp.web.lists.getByTitle("InheritParentPermission").items.add({
                        Title: folderInfo[0].Title,
                        FolderID: iar.data.ID,
                        IsDone: "NO",
                        ParentID: Number(folderInfo[0].ID)
                      });

                      alert(`Dossier ajouté avec succès`);
                    })
                    .then(() => {

                      if (folderInfo[0].FolderID !== 1) {
                        window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].FolderID}`;

                      }

                      else {
                        window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx`;
                      }
                      // window.open("https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=" + subId)
                    });

                });

            }
            catch (err) {
              console.log("Erreur:", err.message);
            }

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

      //
      {

      }

      //azoute permission
      {

        {
          $("#group_name").bind('input', () => {
            const shownVal = (document.getElementById("group_name") as HTMLInputElement).value;
            // var shownVal = document.getElementById("name").value;

            const value2send = (document.querySelector<HTMLSelectElement>(`#groups option[value='${shownVal}']`) as HTMLSelectElement).dataset.value;
            principleIdOfGroup = value2send;
            console.log(value2send);
            //  $("#created_by").val(value2send);
          });
        }
        //add permission user

        var add_user_permission_container: Element = document.getElementById("add_btn_user");

        let add_btn_user_permission: string = `
                <button type="button" class="btn btn-primary add_group mb-2" style="font-size: 1em;" id=${folderInfo[0].ID}_add_user>Ajouter</button>
                `;

        add_user_permission_container.innerHTML = add_btn_user_permission;

        const btn_add_user = document.getElementById(folderInfo[0].ID + '_add_user');

        var peopleID = null;


        await btn_add_user?.addEventListener('click', async () => {


          var selected_permission = $("#permissions_user option:selected").val();

          var permission = 0;

          if ($("#users_name").val() === "") {
            alert("Please select a user.");
          }
          else {

            if (selected_permission === "ALL") {

              permission = 1073741829;
            }

            else if (selected_permission === "READ") {
              permission = 1073741826;

            }
            else if (selected_permission === "READ_WRITE") {
              permission = 1073741830;

            }


            const user: any = await sp.web.siteUsers.getByEmail($("#users_name").val().toString())();

            users_Permission = user;

            console.log("USERS FOR PERMISSION", users_Permission);

            var x = await this.getChildrenById(folderInfo[0].FolderID, []);


            try {
              console.log("KEY", folderInfo[0].FolderID);

              await sp.web.lists.getByTitle("AccessRights").items.add({
                Title: folderInfo[0].Title.toString(),
                groupName: $("#users_name").val(),
                permission: $("#permissions_user option:selected").val(),
                FolderID: folderInfo[0].ID.toString(),
                PrincipleID: user.Id,
                RoleDefID: permission,
                inherit: folderInfo[0].inherit
              })
                .then(async () => {

                  await sp.web.lists.getByTitle("Documents").items.getById(folderInfo[0].ID).update({
                    inheriting: "NO"
                  }).then(result => {
                    console.log("Item updated successfully");
                  }).catch(error => {
                    console.log("Error updating item: ", error);
                  });

                  await Promise.all(x.map(async (item) => {

                    if (item.inheriting !== "NO") {
                      await sp.web.lists.getByTitle("AccessRights").items.add({
                        Title: item.Title.toString(),
                        groupName: $("#users_name").val(),
                        permission: $("#permissions_user option:selected").val(),
                        FolderID: item.ID.toString(),
                        PrincipleID: user.Id,
                        RoleDefID: permission
                      });
                    }
                  }));
                })
                .then(() => {
                  alert("Autorisation ajoutée à ce dossier avec succès.")
                })
                // .then(() => {
                //   sp.web.lists.getByTitle("Documents").items.getById(item.id).update({
                //     inheriting: "NO",
                //   }).then(result => {
                //     console.log("Item updated successfully");
                //   }).catch(error => {
                //     console.log("Error updating item: ", error);
                //   });
                // })
                .then(() => {
                  window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].FolderID}`;
                });

            }

            catch (e) {
              alert("Erreur: " + e.message);
            }

          }

        });

        var add_group_permission_container: Element = document.getElementById("add_btn_group");

        let add_btn_group_permission: string = `
                <button type="button" class="btn btn-primary add_group mb-2" style="font-size: 1em;" id=${folderInfo[0].FolderID}_add_group>Ajouter</button>
                `;

        add_group_permission_container.innerHTML = add_btn_group_permission;

        const btn_add_group = document.getElementById(folderInfo[0].ID + '_add_group');

        await btn_add_group?.addEventListener('click', async () => {

          var selected_permission = $("#permissions_group option:selected").val();

          var permission = 0;



          if ($("#group_name").val() === "") {
            alert("Please select a group.");
          }
          else {

            if (selected_permission === "ALL") {

              permission = 1073741829;
            }

            else if (selected_permission === "READ") {
              permission = 1073741826;

            }
            else if (selected_permission === "READ_WRITE") {
              permission = 1073741830;

            }

            //  const stringGroupUsers: string[] = await getAllUsersInGroup($("#group_name").val());
            //  console.log("TESTER GROUP USERS", stringGroupUsers);

            this.add_permission_group2($("#group_name").val().toString(), permission, folderInfo[0].FolderID, folderInfo[0].Title, folderInfo[0].ID, folderInfo[0].inheriting, principleIdOfGroup);

            await sp.web.lists.getByTitle("Documents").items.getById(folderInfo[0].ID).update({
              inheriting: "NO",
            }).then(result => {
              console.log("Item updated successfully");
            }).catch(error => {
              console.log("Error updating item: ", error);
            });
          }

        });

        var inherit_permission_container: Element = document.getElementById("inheritParentFolderPermission");
        let inherit_parent_permission: string = `
                      <button type="button" class="btn btn-primary add_group mb-2" style="font-size: 1em;" id=${folderInfo[0].ID}_inheritParentPermission>Hériter les droits d'accès du parent</button>
                      `;

        inherit_permission_container.innerHTML = inherit_parent_permission;

        const btn_inherit_permission = document.getElementById(folderInfo[0].ID + '_inheritParentPermission');

        await btn_inherit_permission?.addEventListener('click', async () => {

          var x = await this.getChildrenById(folderInfo[0].FolderID, []);

          try {
            // console.log(item_perm.title);

            var items = await sp.web.lists.getByTitle("Documents").items
              .select("ID")
              .filter(`FolderID eq '${folderInfo[0].ParentID}' and IsFolder eq 'TRUE'`)
              .get();

            await sp.web.lists.getByTitle("InheritParentPermission").items.add({
              Title: folderInfo[0].Title,
              FolderID: folderInfo[0].ID,
              IsDone: "NO",
              ParentID: Number(items[0].ID)
            })
              .then(async () => {
                await Promise.all(x.map(async (item_group) => {
                  await sp.web.lists.getByTitle("InheritParentPermission").items.add({
                    Title: item_group.Title,
                    FolderID: item_group.ID,
                    IsDone: "NO",
                    ParentID: Number(items[0].ID)
                  });
                }));

              })
              .then(() => {
                console.log("ADDED PARENT");
              })
              .then(() => {

                sp.web.lists.getByTitle("Documents").items.getById(folderInfo[0].ID).update({
                  inheriting: "YES",
                }).then(result => {
                  console.log("Item updated successfully");
                }).catch(error => {
                  console.log("Error updating item: ", error);
                });
              });

            alert("Parent permissions added.");
            window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].FolderID}`;

          }
          catch (e) {
            alert(e.message);
          }
        });


      }


      //close doc upload
      {
        $("#cancel_doc").click(() => {

          $("#doc_details_add").css("display", "none");
        });
      }


    }

    //dept_bookmark

    {
      var btn_bookmark: Element = document.getElementById("bouton_bookmark");

      let nav_html_bookmarked: string = '';

      let nav_html_not_bookmarked: string = '';

      nav_html_bookmarked = `
      <a title="Retirer depuis Marque-Pages" 
      role="button" id='${folderInfo[0].ID}_removeBookmark' style="color: rgb(13, 110, 253);">
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512" class="svg-inline--fa fa-bookmark fa-icon fa-2x"><!--! Font Awesome Pro 6.4.0 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license (Commercial License) Copyright 2023 Fonticons, Inc. --><path fill="#ffd700" d="M0 48V487.7C0 501.1 10.9 512 24.3 512c5 0 9.9-1.5 14-4.4L192 400 345.7 507.6c4.1 2.9 9 4.4 14 4.4c13.4 0 24.3-10.9 24.3-24.3V48c0-26.5-21.5-48-48-48H48C21.5 0 0 21.5 0 48z"/></svg>
      </a>`;

      nav_html_not_bookmarked = ` <a title="Ajouter dans Marque-Pages" 
      role="button" id='${folderInfo[0].ID}_addBookmark' style="color: rgb(13, 110, 253);">
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512" class="svg-inline--fa fa-bookmark fa-icon fa-2x"><!--! Font Awesome Pro 6.4.0 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license (Commercial License) Copyright 2023 Fonticons, Inc. --><path fill="#ffd700" d="M0 48C0 21.5 21.5 0 48 0l0 48V441.4l130.1-92.9c8.3-6 19.6-6 27.9 0L336 441.4V48H48V0H336c26.5 0 48 21.5 48 48V488c0 9-5 17.2-13 21.3s-17.6 3.4-24.9-1.8L192 397.5 37.9 507.5c-7.3 5.2-16.9 5.9-24.9 1.8S0 497 0 488V48z"></path></svg>
      </a>`;


      const user = await sp.web.currentUser();
      var items = await sp.web.lists.getByTitle("Marque_Pages").items
        .select("ID")
        .filter(`FolderID eq '${folderInfo[0].FolderID}' and user eq '${user.Title}'`)
        .get();

      if (items.length === 0) {

        btn_bookmark.innerHTML = nav_html_not_bookmarked;

      } else {
        btn_bookmark.innerHTML = nav_html_bookmarked;
      }


      const btn_addBookmark = document.getElementById(folderInfo[0].ID + '_addBookmark');
      const btn_removeBookmark = document.getElementById(folderInfo[0].ID + '_removeBookmark');

      //  var title = document.title;
      let user_current = await sp.web.currentUser();



      await btn_addBookmark?.addEventListener('click', async () => {
        // this.domElement.querySelector('#btn' + item.Id + '_edit').addEventListener('click', () => {
        //localStorage.setItem("contractId", item.Id);


        try {
          await sp.web.lists.getByTitle("Marque_Pages").items.add({
            Title: folderInfo[0].Title,
            url: `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].FolderID}`,
            user: user_current.Title,
            FolderID: folderInfo[0].FolderID
          })
            .then(() => {
              alert("Ajoutee dans Marque-Pages.");
              window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].FolderID}`;
            });

        }
        catch (err) {
          alert(err.message);
        }
      });

      await btn_removeBookmark?.addEventListener('click', async () => {

        try {
          var items = await sp.web.lists.getByTitle("Marque_Pages").items
            .select("ID")
            .filter(`FolderID eq '${folderInfo[0].FolderID}' and user eq '${folderInfo[0].Title}'`)
            .get();

          if (items.length === 0) {
            console.log('Item not found in Favourites list.');
            return;
          }

          // Delete the item from the Favourites list
          await sp.web.lists.getByTitle("Marque_Pages").items.getById(items[0].ID)
            .delete()
            .then(() => {
              alert("Retiree depuis Marque-Pages.");
              window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].FolderID}`;
            });

        }

        catch (e) {
          alert(e.message);
        }


      });

    }

    //dept
    {
      var btn_dept: Element = document.getElementById("ajouterDept");

      let nav_html_dept: string = '';

      let nav_html_not_dept: string = '';

      nav_html_dept = `<a title="Retirer depuis department" id='${folderInfo[0].ID}_removeDept' role="button" style="color: gold;"><svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="square-xmark" class="svg-inline--fa fa-square-xmark fa-icon fa-2x" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512">
    <path fill="currentColor" d="M64 32C28.7 32 0 60.7 0 96V416c0 35.3 28.7 64 64 64H384c35.3 0 64-28.7 64-64V96c0-35.3-28.7-64-64-64H64zm79 143c9.4-9.4 24.6-9.4 33.9 0l47 47 47-47c9.4-9.4 24.6-9.4 33.9 0s9.4 24.6 0 33.9l-47 47 47 47c9.4 9.4 9.4 24.6 0 33.9s-24.6 9.4-33.9 0l-47-47-47 47c-9.4 9.4-24.6 9.4-33.9 0s-9.4-24.6 0-33.9l47-47-47-47c-9.4-9.4-9.4-24.6 0-33.9z"></path></svg></a> `;

      nav_html_not_dept = `<a title="Ajouter dans department" id='${folderInfo[0].ID}_addDept' role="button" style="color: gold;"><svg aria-hidden="true" focusable="false" data-prefix="far" data-icon="square-check" class="svg-inline--fa fa-square-check fa-icon fa-2x" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512">
    <path fill="currentColor" d="M211.8 339.8C200.9 350.7 183.1 350.7 172.2 339.8L108.2 275.8C97.27 264.9 97.27 247.1 108.2 236.2C119.1 225.3 136.9 225.3 147.8 236.2L192 280.4L300.2 172.2C311.1 161.3 328.9 161.3 339.8 172.2C350.7 183.1 350.7 200.9 339.8 211.8L211.8 339.8zM0 96C0 60.65 28.65 32 64 32H384C419.3 32 448 60.65 448 96V416C448 451.3 419.3 480 384 480H64C28.65 480 0 451.3 0 416V96zM48 96V416C48 424.8 55.16 432 64 432H384C392.8 432 400 424.8 400 416V96C400 87.16 392.8 80 384 80H64C55.16 80 48 87.16 48 96z"></path></svg></a>`;


      var items = await sp.web.lists.getByTitle("Department").items
        .select("ID")
        .filter(`FolderID eq '${folderInfo[0].FolderID}'`)
        .get();

      if (items.length === 0) {
        btn_dept.innerHTML = nav_html_not_dept;
      } else {
        btn_dept.innerHTML = nav_html_dept;

      }


      const btn_addDept = document.getElementById(folderInfo[0].ID + '_addDept');
      const btn_removeDept = document.getElementById(folderInfo[0].ID + '_removeDept');

      await btn_addDept?.addEventListener('click', async () => {

        try {

          await sp.web.lists.getByTitle("Department").items.add({
            Title: folderInfo[0].Title,
            url: `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].FolderID}`,
            // user: user_current.Title,
            FolderID: folderInfo[0].FolderID
          })
            .then(() => {
              alert('Dossier ajouté à la liste des départements.');
              window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].FolderID}`;

            });

        }
        catch (e) {

          alert(e.message);
        }


      });

      await btn_removeDept?.addEventListener('click', async () => {

        try {
          var items = await sp.web.lists.getByTitle("Department").items
            .select("ID")
            .filter(`FolderID eq '${folderInfo[0].FolderID}'`)
            .get();

          if (items.length === 0) {
            console.log('Folder not found in Department list.');
            return;
          }

          // Delete the item from the Favourites list
          await sp.web.lists.getByTitle("Department").items.getById(items[0].ID).delete();

          alert('Dossier supprimé de la liste des services.');
          window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].FolderID}`;

        }

        catch (e) {

          alert(e.message);
        }
      });
    }

    const document_container = document.getElementById("tbl_documents_bdy");

    if (!document_container) {
      return;
    }

    document_container.innerHTML = '';

    try {
      const all_documents = await sp.web.lists.getByTitle('Documents').items
        .select("ID,ParentID,FolderID,Title,revision,IsFolder,description,attachmentUrl,IsFiligrane,IsDownloadable, filename")
        .top(5000)
        .filter(`ParentID eq '${itemKey}' and IsFolder eq '${value1}'`)
        .getAll();

      console.log("CLICK LENGTH", all_documents.length);
      console.log("CLICK LENGTH", all_documents);

      response_doc = all_documents;

      const result = response_doc.reduce((acc: any[], obj: any) => {
        if (!obj.revision || obj.revision === null) return acc;
        let existingObj = acc.find(o => o.Title === obj.Title);
        if (!existingObj || obj.revision > existingObj.revision) {
          acc = acc.filter(o => o.Title !== obj.Title);
          acc.push(obj);
        }
        return acc;
      }, []).sort((a: any, b: any) => (a.Title > b.Title) ? 1 : -1);


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
  
            <td class="text-left" style="
            padding-left: 6rem;"> 
            ${element.description}          
            </td>
  
            <td class="text-center">${element.revision}</td>

            
            <td style="font-size: 8px;">
<div class="button-container">
  <a title="Mettre à jour le document" role="button" id="${element.Id}_view_doc_details" class="btn_view_doc_details" style="text-decoration: auto;">
  <svg aria-hidden="true" focusable="false" data-prefix="far" 
  data-icon="eye" class="svg-inline--fa fa-eye fa-icon fa-2x" 
  role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 288 256"><!--! Font Awesome Pro 6.3.0 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license (Commercial License) Copyright 2023 Fonticons, Inc. --><path d="M256 512A256 256 0 1 0 256 0a256 256 0 1 0 0 512zM216 336h24V272H216c-13.3 0-24-10.7-24-24s10.7-24 24-24h48c13.3 0 24 10.7 24 24v88h8c13.3 0 24 10.7 24 24s-10.7 24-24 24H216c-13.3 0-24-10.7-24-24s10.7-24 24-24zm40-208a32 32 0 1 1 0 64 32 32 0 1 1 0-64z"/></svg>
  </a>

  <a title="Voir le document" id="${element.Id}_view_doc" role="button"  class="btn_view_doc">
  <svg aria-hidden="true" focusable="false" data-prefix="far" 
  data-icon="eye" class="svg-inline--fa fa-eye fa-icon fa-2x" 
  role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 288 256">
  <path fill="currentColor" d="M160 256C160 185.3 217.3 128 288 128C358.7 128 416 185.3 416 256C416 326.7 358.7 384 288 384C217.3 384 160 326.7 160 256zM288 336C332.2 336 368 300.2 368 256C368 211.8 332.2 176 288 176C287.3 176 286.7 176 285.1 176C287.3 181.1 288 186.5 288 192C288 227.3 259.3 256 224 256C218.5 256 213.1 255.3 208 253.1C208 254.7 208 255.3 208 255.1C208 300.2 243.8 336 288 336L288 336zM95.42 112.6C142.5 68.84 207.2 32 288 32C368.8 32 433.5 68.84 480.6 112.6C527.4 156 558.7 207.1 573.5 243.7C576.8 251.6 576.8 260.4 573.5 268.3C558.7 304 527.4 355.1 480.6 399.4C433.5 443.2 368.8 480 288 480C207.2 480 142.5 443.2 95.42 399.4C48.62 355.1 17.34 304 2.461 268.3C-.8205 260.4-.8205 251.6 2.461 243.7C17.34 207.1 48.62 156 95.42 112.6V112.6zM288 80C222.8 80 169.2 109.6 128.1 147.7C89.6 183.5 63.02 225.1 49.44 256C63.02 286 89.6 328.5 128.1 364.3C169.2 402.4 222.8 432 288 432C353.2 432 406.8 402.4 447.9 364.3C486.4 328.5 512.1 286 526.6 256C512.1 225.1 486.4 183.5 447.9 147.7C406.8 109.6 353.2 80 288 80V80z">
  </path></svg>
  </a>
</div>
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

                    if (this.getFileExtensionFromUrl(urlFile) !== "pdf") {

                      // if (element.IsFiligrane == "NO") {
                      window.open(`${urlFile}`, '_blank');
                    }

                    else {
                      const blurDiv = document.createElement('div');
                      blurDiv.classList.add('blur');
                      document.body.appendChild(blurDiv);

                      // create a div element to show the loader
                      const loaderDiv = document.createElement('div');
                      loaderDiv.classList.add('loader1');
                      document.body.appendChild(loaderDiv);

                      try {

                        //   await this.openPDFInBrowser(url, 'UNCONTROLLED COPY - Downloaded on: ');
                        //await this.openPDFInIframe(urlFile, 'UNCONTROLLED COPY - Downloaded on: ');
                        // await this.createWebpageInIframe2(urlFile);

                        await this.createWebpageInNewTab(urlFile, element.filename);

                        // await this.createWebpageInNewTab2(urlFile); 
                        //  window.open(`${urlFile}`, '_blank');

                      } finally {
                        // remove the loader and the blur elements
                        document.body.removeChild(loaderDiv);
                        document.body.removeChild(blurDiv);
                      }
                      // }
                      //   window.open(`${urlFile}`, '_blank');
                    }
                  }
                  else {
                    //   window.open(`${externalFileUrl}`, '_blank');
                  }

                });

                //view details_doc
                // await btn_view_doc_details?.addEventListener('click', async () => {
                //   window.open(`https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/Document.aspx?document=${element.Title}&documentId=${element.FolderID}`, '_blank');
                // });

                await btn_view_doc_details?.addEventListener('click', async () => {
                  window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/Document.aspx?document=${element.Title}&documentId=${element.FolderID}`;
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

    }

    catch (error) {
      console.error(error);
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
      left: 50px;
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
        width: 90%;
        height: 90%;
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
                  metaData: { fileName: '${filename}' }
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

  private async getParentID(id: any): Promise<number[]> {

    var parentID: number | null = null;
    var value2 = "TRUE";

    var parentIDArray: number[] = [];

    await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID").filter("FolderID eq '" + id + "' and IsFolder eq '" + value2 + "'").get().then((results) => {
      parentID = results[0].ParentID;
      if (parentID) {
        parentIDArray.push(parentID);
        console.log("Parent 1", parentID);
      }
    });

    while (parentID && parentID != 1) {
      await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID, Title").filter("FolderID eq '" + parentID + "' and IsFolder eq '" + value2 + "'").get().then((results) => {
        parentID = results[0].ParentID;
        if (parentID) {
          if (parentArray) {
            parentIDArray.unshift(parentID);
            console.log("Parent 2", parentID);
          }
        }
      });
    }

    if (id) {
      parentIDArray.push(parseInt(this.getItemId()));
    }

    if (parentIDArray.length > 1) {
      parentIDArray.shift();
    }

    console.log("ArrayParent", parentIDArray);

    return parentIDArray;
  }

  private async addBookmark(docID: any, title: any) {
    // Get the current page URL and title
    var url = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${docID}`;
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

  private async addDept(docID: any, title: any) {
    // Get the current page URL and title
    var url = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${docID}`;
    //  var title = document.title;
    let user_current = await sp.web.currentUser();


    // Add the item to the Favourites list
    await sp.web.lists.getByTitle("Department").items.add({
      Title: title,
      url: url,
      // user: user_current.Title,
      FolderID: docID
    });

    console.log('Folder added to Department list.');
  }

  private async removeDept(docID: any) {
    // Get the current page URL

    // Find the item to delete from the Favourites list
    var items = await sp.web.lists.getByTitle("Department").items
      .select("ID")
      .filter(`FolderID eq '${Number(docID)}'`)
      .get();

    if (items.length === 0) {
      console.log('Folder not found in Department list.');
      return;
    }

    // Delete the item from the Favourites list
    await sp.web.lists.getByTitle("Department").items.getById(items[0].ID).delete();

    console.log('Folder removed from Department list.');
  }


  // public render(): React.ReactElement<IMyGedTreeViewProps> {

  render() {

    const { TreeLinks, parentIDArray, selectedKey, isLoaded } = this.state;

    var y = [];

    x = this.getItemId();
    const defaultSelectedKeys = [x]; // Or whatever keys you want to use

    // this.require_libraries();

    // this.getParentID(x);

    console.log("TEST PARENT ARRAY", y);
    console.log("ITEM TO EXPAND", this.getItemId());
    console.log("BEFORE RENDER", this.state.parentIDArray);


    const handleTreeResponsive = () => {

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
            <img src="https://ncaircalin.sharepoint.com/sites/MyGed/SiteAssets/images/loader.gif" alt="Loading..." />
          </div>
        );
      }
    }
    // this.require_libraries();
    return (

      <div className="container-fluid">

        <div className="row">
          <div className="col-sm-3">
            <div id="sidebarMenu" className="sidebar">

              <div className="close-sidebar" role="button"
                onClick={(event: React.MouseEvent<HTMLElement>) => {
                  // const sidebarMenu = document.querySelector("div:has(>#sidebarMenu)");

                  // sidebarMenu.classList.toggle("sidebar-toggle");

                  const sidebarMenu = document.querySelector("#sidebarMenu")?.parentNode as HTMLElement | null;

                  if (sidebarMenu) {
                    sidebarMenu.classList.toggle("sidebar-toggle");
                  }


                }}>
                <FontAwesomeIcon icon={faCircleXmark} className="fa-icon fa-2x"></FontAwesomeIcon></div>

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
                    // defaultSelectedKeys={defaultSelectedKeys}
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

            <div className='left-arrow-responsive' role="button" onClick={(event: React.MouseEvent<HTMLElement>) => {
              // const sidebarMenu = document.querySelector("div:has(>#sidebarMenu)");
              // sidebarMenu.classList.toggle("sidebar-toggle");

              const sidebarMenu = document.querySelector("#sidebarMenu")?.parentNode as HTMLElement | null;

              if (sidebarMenu) {
                sidebarMenu.classList.toggle("sidebar-toggle");
              }

            }}><FontAwesomeIcon icon={faSquareCaretLeft} className="fa-icon fa-2x"></FontAwesomeIcon></div>

            <div id="loader"></div>

            <form id="form_metadata">

              <div id="access_form">

                <div className="dossier_headers">
                  <div className="container1">
                    <div className="image">
                      <img src='https://ncaircalin.sharepoint.com/sites/MyGed/SiteAssets/images/flower.png' />
                      <h2 id='h2_folderName'>
                        Gestion Documentaire
                      </h2>
                    </div>

                  </div>


                  <nav aria-label="breadcrumb" id='nav'>
                    <ul className="breadcrumb" id="folder_nav">
                      <li className="breadcrumb-item" id="editFolder"><a style={{ color: '#0d6efd' }} title="Mettre à jour le dossier" role="button" onClick={async (event: React.MouseEvent<HTMLElement>) => { await this.load_folders(); $("#access_rights_form").css("display", "none"); $("#alert_0_doc").css("display", "none"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "block"); $("#doc_details_add").css("display", "none"); }}><FontAwesomeIcon icon={faEdit} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                      <li className="breadcrumb-item" id="ajouteDoc"><a style={{ color: '#0d6efd' }} title="Créer un document" role="button" id='add_document' onClick={(event: React.MouseEvent<HTMLElement>) => { $("#access_rights_form").css("display", "none"); $("#alert_0_doc").css("display", "none"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "none"); $("#doc_details_add").css("display", "block"); }}><FontAwesomeIcon icon={faFile} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                      <li className="breadcrumb-item" id="accesFolder"><a style={{ color: '#0d6efd' }} title="Autorisation sur le dossier" role="button"
                        onClick={async (event: React.MouseEvent<HTMLElement>) => {

                          try {
                            // Create the loader element
                            const loader = document.createElement("div");
                            loader.id = "loader2";
                            loader.innerHTML = "<div id='loader-spinner'></div>";
                            document.body.appendChild(loader);

                            await this.getSiteUsers();
                            await this.getSiteGroups();
                            $("#table_documents").css("display", "none"); $("#access_rights_form").css("display", "block"); $("#alert_0_doc").css("display", "none"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "none"); $("#doc_details_add").css("display", "none");


                            // Remove the loader element once the function has finished executing
                            document.body.removeChild(loader);
                          } catch (error) {
                            console.error(error);
                          }

                        }}

                      ><FontAwesomeIcon icon={faLock} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                      <li className="breadcrumb-item" id="addFolder"><a style={{ color: '#0d6efd' }} title="Ajouter des sous-dossiers" role="button" onClick={(event: React.MouseEvent<HTMLElement>) => { $("#access_rights_form").css("display", "none"); $("#subfolders_form").css("display", "block"); $("#edit_details").css("display", "none"); $("#alert_0_doc").css("display", "none"); $("#doc_details_add").css("display", "none"); }}><FontAwesomeIcon icon={faFolderPlus} className="fa-icon fa-2x"></FontAwesomeIcon></a></li>
                      {/* <li className="breadcrumb-item"><a style={{ color: 'gold' }} href="#" title="Ajouter dans marque-pages" role="button" onClick={(event: React.MouseEvent<HTMLElement>) => {
                        this.handleIconClick();
                      }}>
                        <FontAwesomeIcon icon={this.state.isToggledOn ? faSolidBook : faBookmark} className="fa-icon fa-2x" />
                      </a></li> */}


                      <li className="breadcrumb-item" style={{ color: '#0d6efd' }} id='bouton_bookmark'></li>

                      <li className="breadcrumb-item" style={{ color: '#0d6efd' }} id='bouton_delete'></li>
                      {/* <li className="breadcrumb-item" id='bouton_delete'><a href="#" title="Supprimer" role="button" id='delete_folder'><FontAwesomeIcon icon={faTrashCan} className="fa-icon fa-2x"></FontAwesomeIcon></a></li> */}
                      {/* <li className="breadcrumb-item"><a style={{ color: '#0d6efd' }} href="#" title="Notifier" role="button" onClick={async (event: React.MouseEvent<HTMLElement>) => { await this.getSiteUsers(); this.getSiteGroups(); $("#table_documents").css("display", "none"); $("#access_rights_form").css("display", "none"); $("#alert_0_doc").css("display", "none"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "none"); $("#doc_details_add").css("display", "none"); $("#notifications_form").css("display", "block"); }} ><FontAwesomeIcon icon={faBell} className="fa-icon fa-2x"></FontAwesomeIcon></a></li> */}
                      {/* <li className="breadcrumb-item"><a style={{ color: '#0d6efd' }} href="#" title="Notifier" role="button" onClick={async (event: React.MouseEvent<HTMLElement>) => { await this.getSiteUsers(); this.getSiteGroups(); $("#table_documents").css("display", "none"); $("#access_rights_form").css("display", "none"); $("#alert_0_doc").css("display", "none"); $("#subfolders_form").css("display", "none"); $("#edit_details").css("display", "none"); $("#doc_details_add").css("display", "none"); $("#notifications_form").css("display", "block"); }} ><FontAwesomeIcon icon={faBell} className="fa-icon fa-2x"></FontAwesomeIcon></a></li> */}
                      {/* <li className="breadcrumb-item" id="ajouterDept"><a style={{ color: 'gold' }} href="#" title="Ajouter dans department" role="button" onClick={(event: React.MouseEvent<HTMLElement>) => {

                        this.handleIconClickDept();
                      }}>
                        <FontAwesomeIcon icon={this.state.isToggleOnDept ? faSquareXmark : faSquareCheck} className="fa-icon fa-2x" />
                      </a></li> */}

                      <li className="breadcrumb-item" style={{ color: '#0d6efd' }} id='ajouterDept'></li>


                      {/* 
                      <li className="breadcrumb-item" id="ajouterDept">
                        <a style={{ color: 'gold' }} href="#" title="Ajouter dans department" role="button" onClick={(event) => {
                          event.preventDefault();
                          this.handleIconClickDept();
                        }}>
                          <span className="fa-icon fa-2x">
                            <i className="fas" id="deptIcon"></i>
                          </span>
                        </a>
                      </li> */}



                    </ul>
                  </nav>

                </div>


                <h4 id='alert_0_doc'>Veuillez sélectionner un sous répertoire</h4>



                <div id="edit_details">
                  <div className="row">
                    <div className="col-sm-6">
                      <Label>Nom du dossier
                        <input type="text" className="form-control" id="folder_name1" style={{ fontSize: "1em" }} />
                      </Label>
                    </div>

                    <div className="col-sm-6">
                      <Label>Description
                        <textarea className="form-control" id="folder_desc" rows={2} cols={60} style={{ fontSize: "1em" }}></textarea>
                      </Label>
                    </div>
                  </div>



                  <div className="row">
                    <div className="col-sm-6">
                      <Label>Dossier
                        <input type="text" className="form-control" id="parent_folder" list='folders' style={{ fontSize: "1em" }} />

                        <datalist id="folders">
                          <select id="select_folders"></select>
                        </datalist>
                        {/* <select className='form-select' name="parentFolder" id="parent_folder">

            </select> */}
                      </Label>
                    </div>

                    {/* <div className="col-sm-6">
                      <Label>Folder Order
                        <input type="text" className="form-control" id="folder_order" />
                      </Label>
                    </div> */}

                  </div>

                  <div className="row">
                    <div className="col-sm-8">

                    </div>
                    <div className="col-sm-2" id="update_btn_dossier">

                      {/* <button type="button" className="btn btn-primary" id='update_details'>Edit Details</button> */}

                    </div>

                    <div className="col-sm-2">
                      <button type="button" className="btn btn-primary" style={{ fontSize: "1em" }} id='edit_cancel'>Annuler</button>

                    </div>

                  </div>

                </div>

                <div id="access_rights_form">


                  <div className="row">

                    <div className="col-sm-6">
                      <Label>Ajouter un droit d'accès utilisateur

                        <input type="text" className="form-control" id="users_name" list='users' style={{ fontSize: "1em" }} />

                        <datalist id="users">
                          <select id="select_users"></select>
                        </datalist>
                      </Label>
                    </div>


                    <div className="col-sm-3">
                      <Label style={{
                        display: "flex",
                        flexDirection: "column"
                      }}> Type
                        <select className='form-select' name="permissions" id="permissions_user" style={{ fontSize: "1em" }}>
                          {/* <option value="NONE">NONE</option> */}
                          <option value="READ">READ</option>
                          <option value="READ_WRITE">READ_WRITE</option>
                          <option value="ALL">ALL</option>
                        </select>
                      </Label>
                    </div>
                    <div className="col-sm-3" id="add_btn_user">
                    </div>
                  </div>

                  <div className="row">


                    <div className="col-sm-6">
                      <Label>Ajouter un droit d'accès de groupe
                        <input type="text" className="form-control" id="group_name" list='groups' style={{ fontSize: "1em" }} />

                        <datalist id="groups">
                          {/* <select id="select_groups"></select> */}
                        </datalist>
                      </Label>
                    </div>


                    <div className="col-sm-3">
                      <Label style={{
                        display: "flex",
                        flexDirection: "column"
                      }}
                      > Type
                        <select className='form-select' name="permissions" id="permissions_group" style={{ fontSize: "1em" }}>
                          {/* <option value="NONE">NONE</option> */}
                          <option value="READ">READ</option>
                          <option value="READ_WRITE">READ_WRITE</option>
                          <option value="ALL">ALL</option>
                        </select>
                      </Label>
                    </div>
                    <div className="col-sm-3" id="add_btn_group">
                    </div>
                  </div>


                  <div className="row">
                    <div className='col-sm-6'>
                      <p id="inheritparagraph" className='h4' style={{
                        display: 'none',
                      }}>Ce dossier hérite des permissions de son parent.</p>
                    </div>
                  </div>


                  <div className="row">
                    <div className="col-sm-3" id="inheritParentFolderPermission" >

                    </div>

                  </div>

                  <div className='row'>
                    <div id="spListPermissions" style={{
                      margin: '2em',
                      height: '100%',
                      padding: '2em',
                      boxShadow: '0 4px 8px 0 rgba(0,0,0,.2), 0 6px 20px 0 rgba(0,0,0,.19)',
                      marginTop: '2em',
                      backgroundColor: 'snow'
                    }}>

                    </div>
                  </div>

                </div>

                <div id="notifications_form">


                  <div className="row">

                    <div className="col-sm-6">
                      <Label>Ajouter une notification utilisateur :

                        <input type="text" className="form-control" id="users_name_notif" list='users' style={{ fontSize: "1em" }} />

                        <datalist id="users">
                          <select id="select_users" style={{ fontSize: "1em" }} ></select>
                        </datalist>
                      </Label>
                    </div>


                    <div className="col-sm-3" id="add_btn_user_notif">
                    </div>
                  </div>

                  <div className="row">


                    <div className="col-sm-6">
                      <Label>Ajouter une notification de groupe :
                        <input type="text" className="form-control" id="group_name_notif" list='groups' style={{ fontSize: "1em" }} />

                        <datalist id="groups">
                          {/* <select id="select_groups" style={{ fontSize: "1em" }} ></select> */}
                        </datalist>
                      </Label>
                    </div>



                    <div className="col-sm-3" id="add_btn_group_notif">
                    </div>
                  </div>



                  <div className='row'>
                    <div id="spListNotification">
                      <table id='tbl_notif' className='table table-striped'>
                        <thead>
                          <tr>
                            <th className="text-left">Nom</th>
                            <th className="text-left" >Droits d'accès</th>
                            <th className="text-left" >Actions</th>
                          </tr>
                        </thead>
                        <tbody id="tbl_notif_bdy">



                        </tbody>
                      </table>
                    </div>
                  </div>

                </div>

                <div id='table_documents'>

                  <div id="spListDocuments1">

                    <table id='tbl_documents' className='table1 table-striped'>
                      <thead>
                        <tr>
                          <th className="text-left" id='nom_doc'>Nom du document</th>
                          <th className="text-left" id='desc_doc'>Description</th>
                          <th className="text-center">Revision</th>
                          <th className="text-center" >Actions</th>
                        </tr>
                      </thead>
                      <tbody id="tbl_documents_bdy">


                      </tbody>
                    </table>
                  </div>


                </div>

                <div id="subfolders_form">
                  <div className="row">
                    <div className="col-sm-6">
                      <Label>Nom du Dossier
                        <input type="text" className="form-control" id="folder_name" style={{ fontSize: "1em" }} />
                      </Label>
                    </div>

                    <div className="col-sm-3" id="add_btn_subFolder">

                    </div>
                    <div className="col-sm-3">
                      <button type="button" className="btn btn-primary add_subfolder mb-2 " id="cancel_add_sub" style={{ fontSize: "1em" }} >Annuler</button>
                    </div>
                  </div>

                </div>

                <div id="doc_details_add">


                  <div className="row">
                    <div className="col-sm-6">
                      <Label>Nom du document
                        <input type="text" id='input_doc_number_add' className='form-control' required style={{ fontSize: "1em" }} />
                      </Label>
                    </div>

                    <div className="col-sm-6">
                      <Label>Fichier
                        <input type="file" name="file" id="file_ammendment" className="form-control" style={{ fontSize: "1em" }} />
                      </Label>
                    </div>


                  </div>

                  <div className="row">
                    <div className="col-sm-6">
                      {/* <Label>
                        <input type="checkbox" name="checkFiligrane" className="form-check-input" style={{ fontSize: "1em" }} />
                        Ajouter un filigrane sur le document ?
                      </Label> */}

                      <div className="form-check" style={{ paddingLeft: "0.6em" }}>
                        <input
                          id="watermark-checkbox"
                          className="form-check-input"
                          type="checkbox"
                          name="checkFiligrane"
                          style={{
                            fontSize: "1em", width: "1.5em",
                            height: "1.5em"
                          }}
                        />
                        <label
                          htmlFor="watermark-checkbox"
                          className="form-check-label"
                        >
                          Ajouter un filigrane sur le document
                        </label>
                      </div>
                    </div>

                    <div className="col-sm-6">
                      {/* <Label>
                        <input type="checkbox" name="checkImprimab" className="form-check-input" style={{ fontSize: "1em" }} /> Document imprimable
                      </Label> */}

                      <div className="form-check" style={{ paddingLeft: "0.6em" }}>
                        <input
                          id="printable-checkbox"
                          className="form-check-input"
                          type="checkbox"
                          name="checkImprimab"
                          style={{
                            fontSize: "1em", width: "1.5em",
                            height: "1.5em"
                          }}
                        />
                        <label
                          htmlFor="printable-checkbox"
                          className="form-check-label"
                        >
                          Document imprimable
                        </label>
                      </div>
                    </div>
                  </div>


                  <div className="row">
                    <div className="col-sm-6">
                      <Label>
                        Révision
                        <input type="text" id='input_revision_add' className='form-control' style={{ fontSize: "1em" }} />
                      </Label>
                    </div>
                    <div className="col-sm-6">
                      <Label>
                        Nom de fichier
                        <input type="text" id='input_filename_add' className='form-control' disabled style={{ fontSize: "1em" }} />
                      </Label>
                    </div>
                    {/* <div className="col-sm-3">
                      <Label>
                        Status
                        <input type="text" id='input_status_add' className='form-control' style={{ fontSize: "1em" }} />
                      </Label>
                    </div>
                    <div className="col-sm-3">
                      <Label>
                        Owner
                        <input type="text" id='input_owner_add' className='form-control' style={{ fontSize: "1em" }} />
                      </Label>
                    </div>
                    <div className="col-sm-3">
                      <Label>
                        Active Date
                        <input type="text" id='input_activeDate_add' className='form-control' style={{ fontSize: "1em" }} />
                      </Label>
                    </div> */}
                  </div>

                  {/* <div className="row">
                    <div className="col-sm-6">
                      <Label>
                        Filename
                        <input type="text" id='input_filename_add' className='form-control' disabled style={{ fontSize: "1em" }} />
                      </Label>
                    </div>
                    {/* <div className="col-sm-6">
                      <Label>
                        Author
                        <input type="text" id='input_author_add' className='form-control' style={{ fontSize: "1em" }} />
                      </Label>
                    </div> */}

                  {/* </div>  */}

                  <div className="row">
                    <div className="col-sm-6">
                      <Label>
                        Description
                        <textarea id='input_description_add' className='form-control' rows={2} style={{ fontSize: "1em" }} />
                      </Label>
                    </div>
                    <div className="col-sm-6">
                      <Label>
                        Mots-clés
                        <textarea id='input_keywords_add' className='form-control' rows={2} style={{ fontSize: "1em" }} />
                      </Label>
                    </div>
                    {/* <div className="col-sm-3">
                      <Label>
                        Review Date
                        <input type="text" id='input_reviewDate_add' className='form-control' style={{ fontSize: "1em" }} />
                      </Label>
                    </div> */}
                  </div>

                  <div className="row">
                    <div className="col-sm-8">

                    </div>
                    <div className="col-sm-2" id="add_document_btn">



                    </div>

                    <div className="col-sm-2">
                      <button type="button" className="btn btn-primary" id='cancel_doc' style={{ fontSize: "1em" }} >Annuler</button>
                    </div>

                  </div>

                </div>

              </div>

            </form>

          </div>

        </div >



      </div >
    );

  }

  private async load_folders() {


    var value1 = "TRUE";
    var drp_folders = document.getElementById("select_folders");

    // const allItems: any = await sp.web.lists.getByTitle('Documents').items.getAll(),

    const all_folders: any = await sp.web.lists.getByTitle('Documents').items.select("ID,ParentID,FolderID,Title,IsFolder,description").top(5000).filter("IsFolder eq '" + value1 + "'").get();

    folders = all_folders;

    folders.forEach((result: any) => {
      var opt = document.createElement('option');
      opt.appendChild(document.createTextNode(result.Title));
      opt.value = result.FolderID + "_" + result.Title;
      drp_folders.appendChild(opt);

    });

  }

  public async getSiteUsers() {

    var drp_users = document.getElementById("select_users");
    drp_users.innerHTML = "";


    const users1: any = await sp.web.siteUsers();

    users = users1;

    users.forEach((result: ISiteUserInfo) => {

      if (result.UserPrincipalName != null) {

        console.log("USER", result.Id, result.Email);
        var opt = document.createElement('option');
        opt.appendChild(document.createTextNode(result.Email));
        opt.value = result.Email;
        drp_users.appendChild(opt);
      }

    });

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

  public async getSiteGroups() {

    //var drp_users = document.getElementById("select_groups");
    var drp_users = document.getElementById("groups") as HTMLSelectElement;

    if (!drp_users) {
      console.error("Dropdown element not found");
      return;
    }


    try {

      const { permissions } = await this.getSitePermissions('https://ncaircalin.sharepoint.com/sites/MyGed', "mgolapkhan.ext@aircalin.nc", "musharaf2897");

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

  private async addSubfolders(item: ITreeItem) {

    console.log("ID", item.id);
  }

  private async onTreeItemSelect(items: ITreeItem[]) {

    items.forEach((item: any) => {
      $("#h2_folderName").text(item.label);
    });

  }

  private getDossierTitle() {
    var queryParms = new URLSearchParams(document.location.search.substring(1));
    var myParm = queryParms.get("title");
    if (myParm) {
      return myParm.trim();
    }
  }

  private getDossierID() {
    var queryParms = new URLSearchParams(document.location.search.substring(1));
    var myParm = queryParms.get("folderID");
    if (myParm) {
      return myParm.trim();
    }
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

  private handleDeleteButtonClick(event: MouseEvent, key: any, label: any, name: any, pID: any): void {
    const button = event.target as HTMLButtonElement;
    const id = button.id.slice(0, -5); // remove '_edit' from the button id to get the group id

    try {
      //  console.log("KEY", item.key);

      sp.web.lists.getByTitle("AccessRights").items.add({
        // Title: item.label.toString(),
        Title: label.toString(),

        //  groupName: $(`${element.ID}_personName`).val(),
        groupName: name,
        // groupName: "zpeerbaccus.ext@aircalin.nc",
        permission: "NONE",
        FolderID: Number(key),
        // PrincipleID: user.Id
        // PrincipleID: 15
        PrincipleID: pID


      })
        .then(() => {
          alert("Autorisation supprimer à ce dossier avec succès.");
        })
        .then(() => {
          window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${key}`;
        });
    }

    catch (e) {
      alert("Erreur: " + e.message);
    }
    // Here, you can write code to handle the delete button click, for example:
    // 1. Make an API call to delete the group with the specified id
    // 2. Remove the table row corresponding to the deleted group from the DOM
    // 3. Show a success message to the user, etc.

    console.log(`Delete button clicked for group with id: ${id}`);
  }

  private toggleIcon = async (xx: any) => {
    const user = await sp.web.currentUser();
    var items = await sp.web.lists.getByTitle("Marque_Pages").items
      .select("ID")
      .filter(`FolderID eq '${xx}' and user eq '${user.Title}'`)
      .get();

    if (items.length === 0) {
      this.setState({ isToggledOn: false });
    } else {
      this.setState({ isToggledOn: true });
    }

  };

  private toggleIconDept = async (xx: any) => {
    var items = await sp.web.lists.getByTitle("Department").items
      .select("ID")
      .filter(`FolderID eq '${xx}'`)
      .get();

    if (items.length === 0) {
      this.setState({ isToggleOnDept: false });
    } else {
      this.setState({ isToggleOnDept: true });
    }

  };

  private checkIcons = async (itemkey: any): Promise<void> => {
    var items = await sp.web.lists.getByTitle("Department").items
      .select("ID")
      .filter(`FolderID eq '${itemkey}'`)
      .get();

    // if (this.state.isToggleOnDept === undefined) {
    //   this.setState({ isToggleOnDept: false });
    // }

    if (items.length === 0) {
      this.setState({ isToggleOnDept: true });
    } else {
      this.setState({ isToggleOnDept: false });
    }
  }

  private renderCustomTreeItem(item: ITreeItem): JSX.Element {



    const checkIfUserIsAdmin = async (graphClient: MSGraphClient): Promise<boolean> => {
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

    const checkIfUserIsRefUser = async (graphClient: MSGraphClient): Promise<boolean> => {
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

    const checkIfUserIsGuestUser = async (graphClient: MSGraphClient): Promise<boolean> => {
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

    const generatePdfBytes = async (fileUrl: string, filigraneText: string): Promise<Uint8Array> => {
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

    const openPDFInObject = async (url) => {
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

      const loader = document.createElement('div');
      loader.style.cssText = `
        display: flex;
        justify-content: center;
        align-items: center;
        width: 100%;
        height: 100%;
        position: absolute;
        top: 0;
        left: 0;
      `;

      const loaderText = document.createElement('div');
      loaderText.style.cssText = `
        font-size: 24px;
        color: #fff;
      `;
      loaderText.innerText = 'Loading...';
      loader.appendChild(loaderText);

      overlay.appendChild(loader);

      // const pdfBytes = await this.generatePdfBytes(url, filigraneText);
      // const pdfUrl = URL.createObjectURL(new Blob([pdfBytes], { type: 'application/pdf' }));

      const iframe = document.createElement('iframe');
      iframe.src = `${url}#toolbar=0`;
      iframe.style.cssText = `
        border: none;
        width: 100%;
        height: 100%;
        max-width: 1000px;
        max-height: 90vh;
      `;
      // iframe.setAttribute('sandbox', 'allow-same-origin allow-popups allow-scripts');

      iframe.addEventListener('load', () => {
        loader.style.display = 'none';
      });

      iframe.addEventListener('contextmenu', (event) => {
        event.preventDefault();
      });

      const closeButton = document.createElement('button');
      closeButton.innerText = 'Close';
      closeButton.style.cssText = `
        position: absolute;
        top: 20px;
        right: 20px;
        background-color: #000;
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

    const loadScript = (src) => {
      return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = src;
        script.onload = resolve;
        script.onerror = reject;
        document.head.appendChild(script);
      });
    };

    const viewPdf = async (pdfUrl) => {
      // Create a new canvas element to display the PDF
      const canvas = document.createElement('canvas');
      canvas.style.display = 'block';
      canvas.style.width = '100%';
      canvas.style.height = '100%';
      document.body.appendChild(canvas);


      await Promise.all([
        loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.8.335/pdf.js'),
        loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.8.335/pdf.worker.js')
      ]);

      // Use PDF.js to load and display the PDF
      try {
        const loadingTask = pdfjsLib.getDocument(pdfUrl);
        const pdf = await loadingTask.promise;

        const page = await pdf.getPage(1);
        const viewport = page.getViewport({ scale: 1 });

        canvas.height = viewport.height;
        canvas.width = viewport.width;

        const renderContext = {
          canvasContext: canvas.getContext('2d'),
          viewport: viewport,
        };
        await page.render(renderContext).promise;
      } catch (error) {
        console.error('Error loading PDF:', error);
      }
    }

    const createWebpageInNewTab = async (url, filename: any) => {
      try {

        const pdfBytes = await generatePdfBytes2(url);
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
        left: 50px;
        padding: 5px 10px;
        background-color: #fa0f00;
        border: none;
        cursor: pointer;
        font-size: large;
        color: white;
        `;

        //       const mediaQuery = window.matchMedia('(max-width: 600px)');
        //       if (mediaQuery.matches) {
        //         closeButton.style.cssText += `
        // position: absolute;
        // bottom: 3px;
        // right: 19px;
        // padding: 5px 10px;
        // background-color: darkblue;
        // border: none;
        // cursor: pointer;
        // font-size: smaller;
        // color: white;
        // `;
        //       }

        overlay.appendChild(closeButton);

        const iframe = document.createElement('iframe');
        iframe.style.cssText = `
          width: 90%;
          height: 90%;
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
                    metaData: { fileName: '${filename}' }
                  }, { embedMode: "IN_LINE", showDownloadPDF: false, showPrintPDF: false });
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

    const add_permission_group2 = async (group_name: string, permission: any, id: any, principleId: any) => {

      try {

        var x = await getChildrenById(id, []);

        await sp.web.lists.getByTitle("AccessRights").items.add({
          Title: item.label.toString(),
          groupName: group_name,
          permission: $("#permissions_group option:selected").val(),
          FolderID: item.id,
          PrincipleID: principleId,
          groupTitle: group_name,
          RoleDefID: permission
        })

          .then(async () => {

            await Promise.all(x.map(async (item_group) => {

              await sp.web.lists.getByTitle("AccessRights").items.add({
                Title: item_group.Title.toString(),
                groupName: group_name,
                permission: $("#permissions_group option:selected").val(),
                FolderID: item_group.ID,
                PrincipleID: principleId,
                groupTitle: group_name,
                RoleDefID: permission
              });

            }));

          });
        // }));

        alert("Authorization added successfully.");
        window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${item.key}`;
      }
      catch (e) {
        alert("Error: " + e.message);
      }
    }

    const getCheckboxValue = (checkbox: HTMLInputElement): string => {
      return checkbox.checked ? "YES" : "NO";
    }

    const getFileExtensionFromUrl = (url: string): string => {
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

    const fetchDocuments = async (itemKey: number): Promise<void> => {
      let response_doc: any = null;
      let response_distinc: any[] = [];
      let html_document = '';
      let value1 = "FALSE";

      let pdfName = '';

      console.log("ITEM KEY", itemKey);


      const document_container = document.getElementById("tbl_documents_bdy");

      if (!document_container) {
        return;
      }

      document_container.innerHTML = '';

      try {
        const all_documents = await sp.web.lists.getByTitle('Documents').items
          .select("ID,ParentID,FolderID,Title,revision,IsFolder,description,attachmentUrl,IsFiligrane,IsDownloadable,inheriting,filename")
          .top(5000)
          .filter(`ParentID eq '${itemKey}' and IsFolder eq '${value1}'`)
          .getAll();

        // if (all_documents[0].inheriting === "YES") {
        //   $("#inheritparagraph").css("display", "block");
        // }

        // else {
        //   $("#inheritparagraph").css("display", "none");
        // }


        console.log("CLICK LENGTH", all_documents.length);
        console.log("CLICK LENGTH", all_documents);

        response_doc = all_documents;

        const result = response_doc.reduce((acc: any[], obj: any) => {
          if (!obj.revision || obj.revision === null) return acc;
          let existingObj = acc.find(o => o.Title === obj.Title);
          if (!existingObj || obj.revision > existingObj.revision) {
            acc = acc.filter(o => o.Title !== obj.Title);
            acc.push(obj);
          }
          return acc;
        }, []).sort((a: any, b: any) => (a.Title > b.Title) ? 1 : -1);


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
    
              <td class="text-center">${element.revision}</td>
  
              <td style="font-size: 8px;">
  <div class="button-container">
    <a title="Mettre à jour le document" role="button" id="${element.Id}_view_doc_details" class="btn_view_doc_details" style="text-decoration: auto;">
    <svg aria-hidden="true" focusable="false" data-prefix="far" 
    data-icon="eye" class="svg-inline--fa fa-eye fa-icon fa-2x" 
    role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 288 256"><!--! Font Awesome Pro 6.3.0 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license (Commercial License) Copyright 2023 Fonticons, Inc. --><path d="M256 512A256 256 0 1 0 256 0a256 256 0 1 0 0 512zM216 336h24V272H216c-13.3 0-24-10.7-24-24s10.7-24 24-24h48c13.3 0 24 10.7 24 24v88h8c13.3 0 24 10.7 24 24s-10.7 24-24 24H216c-13.3 0-24-10.7-24-24s10.7-24 24-24zm40-208a32 32 0 1 1 0 64 32 32 0 1 1 0-64z"/></svg>
    </a>

    <a title="Voir le document" id="${element.Id}_view_doc" role="button"  class="btn_view_doc">
    <svg aria-hidden="true" focusable="false" data-prefix="far" 
    data-icon="eye" class="svg-inline--fa fa-eye fa-icon fa-2x" 
    role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 288 256">
    <path fill="currentColor" d="M160 256C160 185.3 217.3 128 288 128C358.7 128 416 185.3 416 256C416 326.7 358.7 384 288 384C217.3 384 160 326.7 160 256zM288 336C332.2 336 368 300.2 368 256C368 211.8 332.2 176 288 176C287.3 176 286.7 176 285.1 176C287.3 181.1 288 186.5 288 192C288 227.3 259.3 256 224 256C218.5 256 213.1 255.3 208 253.1C208 254.7 208 255.3 208 255.1C208 300.2 243.8 336 288 336L288 336zM95.42 112.6C142.5 68.84 207.2 32 288 32C368.8 32 433.5 68.84 480.6 112.6C527.4 156 558.7 207.1 573.5 243.7C576.8 251.6 576.8 260.4 573.5 268.3C558.7 304 527.4 355.1 480.6 399.4C433.5 443.2 368.8 480 288 480C207.2 480 142.5 443.2 95.42 399.4C48.62 355.1 17.34 304 2.461 268.3C-.8205 260.4-.8205 251.6 2.461 243.7C17.34 207.1 48.62 156 95.42 112.6V112.6zM288 80C222.8 80 169.2 109.6 128.1 147.7C89.6 183.5 63.02 225.1 49.44 256C63.02 286 89.6 328.5 128.1 364.3C169.2 402.4 222.8 432 288 432C353.2 432 406.8 402.4 447.9 364.3C486.4 328.5 512.1 286 526.6 256C512.1 225.1 486.4 183.5 447.9 147.7C406.8 109.6 353.2 80 288 80V80z">
    </path></svg>
    </a>
  </div>
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

                      if (getFileExtensionFromUrl(urlFile) !== "pdf") {
                        // if (element.IsFiligrane == "NO") {
                        window.open(`${urlFile}`, '_blank');
                      }

                      else {
                        const blurDiv = document.createElement('div');
                        blurDiv.classList.add('blur');
                        document.body.appendChild(blurDiv);

                        // create a div element to show the loader
                        const loaderDiv = document.createElement('div');
                        loaderDiv.classList.add('loader1');
                        document.body.appendChild(loaderDiv);

                        try {
                          //   await this.openPDFInBrowser(url, 'UNCONTROLLED COPY - Downloaded on: ');
                          //await openPDFInIframe(urlFile, 'UNCONTROLLED COPY - Downloaded on: ');
                          //createWebpageInNewTab2
                          // await openPDFInEmbed(urlFile);
                          //   await openPDFInObject(urlFile);
                          //await createWebpageInIframe2(urlFile);
                          //  await createWebpageInNewTab2(urlFile);
                          //await createWebpageInNewTab10(urlFile);
                          // viewPdf(urlFile);

                          await createWebpageInNewTab(urlFile, element.filename);

                          //  window.open(`${urlFile}`, '_blank');

                          // loadViewer(urlFile);
                          //openPdf(urlFile);
                          //  btn_view_doc

                        } finally {
                          // remove the loader and the blur elements
                          document.body.removeChild(loaderDiv);
                          document.body.removeChild(blurDiv);

                        }
                        // }
                        //   window.open(`${urlFile}`, '_blank');
                      }

                    }
                    else {
                      window.open(`${externalFileUrl}`, '_blank');
                    }
                  });

                  //view details_doc
                  // await btn_view_doc_details?.addEventListener('click', async () => {
                  //   window.open(`https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/Document.aspx?document=${element.Title}&documentId=${element.FolderID}`, '_blank');
                  // });

                  await btn_view_doc_details?.addEventListener('click', async () => {
                    window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/Document.aspx?document=${element.Title}&documentId=${element.FolderID}`;
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
      }

      catch (error) {
        console.error(error);
      }

    }

    const generatePdfBytes2 = async (fileUrl: string): Promise<Uint8Array> => {
      try {
        const existingPdfBytes = await fetch(fileUrl).then(res => res.arrayBuffer());
        return new Uint8Array(existingPdfBytes);
      } catch (e) {
        console.error('Failed to generate PDF bytes:', e);
        throw e;
      }
    }

    const createWebpageInIframe2 = async (url) => {

      const pdfBytes = await generatePdfBytes2(url);
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
        top: 87px;
        right: 130px;
        padding: 5px 10px;
        background-color: white;
        border: none;
        cursor: pointer;
        font-size: large;
        color: black;
      `;

      overlay.appendChild(closeButton);

      const iframe = document.createElement('iframe');
      iframe.style.cssText = `
        width: 50%;
        height: 98%;
      `;
      iframe.src = 'about:blank';

      overlay.appendChild(iframe);
      document.body.appendChild(overlay);

      closeButton.addEventListener('click', function () {
        document.body.removeChild(overlay);
      });

      const iframeDoc = iframe.contentDocument || iframe.contentWindow.document;

      const htmlContent = `
  <html>
    <head>
      <title>MyGed Viewer</title>
      <style>
        body { margin: 0; }
        canvas { display: block; }
        .pdfjs-toolbar {
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          background-color: #f2f2f2;
          padding: 8px;
          z-index: 9999;
        }
  
        #loader {
          display: none;
          position: fixed;
          top: 50%;
          left: 50%;
          transform: translate(-50%, -50%);
          background-color: #f2f2f2;
          padding: 16px;
          border-radius: 4px;
        }
      </style>
      <script src="https://cdn.jsdelivr.net/npm/pdfjs-dist@2.12.313/build/pdf.js"></script>
      <script>
  
  
        const url = '${url}';
  
        function renderPDF() {
  
          const loader = document.getElementById("loader");
    loader.style.display = "block";
  
          const loadingTask = pdfjsLib.getDocument({ url, enableToolbar: true });
          loadingTask.promise.then(function(pdf) {
            const numPages = pdf.numPages;
            const container = document.createElement("div");
            document.body.appendChild(container);
  
          
  
            function renderPage(pageNumber) {
              pdf.getPage(pageNumber).then(function(page) {
                const viewport = page.getViewport({ scale: 5.3});
                const canvas = document.createElement("canvas");
                const context = canvas.getContext("2d");
                canvas.width = viewport.width;
                canvas.height = viewport.height;
  
                canvas.style.width = "100%"; 
                canvas.style.height = "auto";
                container.appendChild(canvas);
  
                page.render({ canvasContext: context, viewport: viewport }).promise.then(function() {
                  if (pageNumber < numPages) {
                    renderPage(pageNumber + 1); // Render the next page
                  }
                });
              });
            }
  
            renderPage(1);
            loader.style.display = "none"; // Start rendering from the first page
          });
        }
  
  
  
        window.addEventListener("DOMContentLoaded", function() {
  
          pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@2.12.313/build/pdf.worker.min.js";
          pdfRenderOptions = {
            // where cmaps are downloaded from
            cMapUrl: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@2.12.313/cmaps/',
            // The cmaps are compressed in the case
            cMapPacked: true,
            // any other options for pdfjsLib.getDocument.
            // params: {}
          }
  
  
        //  pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.11.338/pdf.worker.min.js";
          pdfjsLib.getDocument(url).promise.then(renderPDF);
        });
  
  
  
        document.addEventListener("contextmenu", function(event) {
          event.preventDefault();
        });
      </script>
    </head>
    <body>
    <div id="loader">Loading...</div>
    
    </body>
  </html>
  `;


      iframeDoc.open();
      iframeDoc.write(htmlContent);
      iframeDoc.close();
    }

    const displayMetadata = (label: any) => {

      {
        $("#access_form").css("display", "block");
        $("#doc_form").css("display", "none");
        $(".dossier_headers").css("display", "block");

        $("#subfolders_form").css("display", "none");

        $("#access_rights_form").css("display", "none");
        $("#notifications_doc_form").css("display", "none");

        $("#doc_details_add").css("display", "none");
        $("#edit_details").css("display", "none");
        $("#h2_folderName").text(label);
      }

      $("#h2_folderName").text(label);
    }

    const deleteDossier = async (id: any, label: any, parentID: any): Promise<void> => {

      var delete_dossier: Element = document.getElementById("bouton_delete");


      let nav_html_delete_dossier: string = '';


      // console.log("ONSELECT", label);

      nav_html_delete_dossier = `
                    <a title="Archiver" 
                    role="button" id='${id}_deleteFolder' style="color: rgb(13, 110, 253);">
                    <img src="https://icons.iconarchive.com/icons/fa-team/fontawesome/128/FontAwesome-Box-Archive-icon.png" id="archiver_dossier" width="34" height="34" style="margin-top: -16px;"></img>
                    </a>`;

      delete_dossier.innerHTML = nav_html_delete_dossier;

      const btn = document.getElementById(id + '_deleteFolder');

      await btn?.addEventListener('click', async () => {
        // this.domElement.querySelector('#btn' + Id + '_edit').addEventListener('click', () => {
        //localStorage.set"contractId", Id);
        // if (confirm(`Êtes-vous sûr de vouloir supprimer ${label} ?`)) {

        //   try {
        //     var res = await sp.web.lists.getByTitle('Documents').items.getById(parseInt(id)).delete()
        //       .then(() => {
        //         alert("Dossier supprimé avec succès.");
        //       })
        //       .then(() => {
        //         window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx`;
        //       });
        //   }
        //   catch (err) {
        //     alert(err.message);
        //   }

        // }
        // else {

        // }

        if (confirm(`Voulez-vous vraiment archiver ce dossier : ${label} ?`)) {

          try {

            const list = sp.web.lists.getByTitle("Documents");

            const i = await list.items.getById(Number(id)).update({
              ParentID: 791,
            })
              .then(() => {
                alert("Dossier archivé avec succès.");
              })
              .then(() => {
                window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx`;
                // window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].ParentID}`;
              });


            // var res = await sp.web.lists.getByTitle('Documents').items.getById(parseInt(folderInfo[0].ID)).delete()
            //   .then(() => {
            //     alert("Dossier supprimé avec succès.");
            //   })
            //   .then(() => {
            //     window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx`;

            //     // window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].ParentID}`;
            //   });
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

    const handleBookmark = async () => {

      var btn_bookmark: Element = document.getElementById("bouton_bookmark");

      let nav_html_bookmarked: string = '';

      let nav_html_not_bookmarked: string = '';

      nav_html_bookmarked = `
      <a title="Retirer depuis Marque-Pages" 
      role="button" id='${item.id}_removeBookmark' style="color: rgb(13, 110, 253);">
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512" class="svg-inline--fa fa-bookmark fa-icon fa-2x"><!--! Font Awesome Pro 6.4.0 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license (Commercial License) Copyright 2023 Fonticons, Inc. --><path fill="#ffd700" d="M0 48V487.7C0 501.1 10.9 512 24.3 512c5 0 9.9-1.5 14-4.4L192 400 345.7 507.6c4.1 2.9 9 4.4 14 4.4c13.4 0 24.3-10.9 24.3-24.3V48c0-26.5-21.5-48-48-48H48C21.5 0 0 21.5 0 48z"/></svg>
      </a>`;

      nav_html_not_bookmarked = ` <a title="Ajouter dans Marque-Pages" 
      role="button" id='${item.id}_addBookmark' style="color: rgb(13, 110, 253);">
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512" class="svg-inline--fa fa-bookmark fa-icon fa-2x"><!--! Font Awesome Pro 6.4.0 by @fontawesome - https://fontawesome.com License - https://fontawesome.com/license (Commercial License) Copyright 2023 Fonticons, Inc. --><path fill="#ffd700" d="M0 48C0 21.5 21.5 0 48 0l0 48V441.4l130.1-92.9c8.3-6 19.6-6 27.9 0L336 441.4V48H48V0H336c26.5 0 48 21.5 48 48V488c0 9-5 17.2-13 21.3s-17.6 3.4-24.9-1.8L192 397.5 37.9 507.5c-7.3 5.2-16.9 5.9-24.9 1.8S0 497 0 488V48z"></path></svg>
      </a>`;


      const user = await sp.web.currentUser();
      var items = await sp.web.lists.getByTitle("Marque_Pages").items
        .select("ID")
        .filter(`FolderID eq '${item.key}' and user eq '${user.Title}'`)
        .get();

      if (items.length === 0) {

        btn_bookmark.innerHTML = nav_html_not_bookmarked;

      } else {
        btn_bookmark.innerHTML = nav_html_bookmarked;
      }


      const btn_addBookmark = document.getElementById(item.id + '_addBookmark');
      const btn_removeBookmark = document.getElementById(item.id + '_removeBookmark');

      //  var title = document.title;
      let user_current = await sp.web.currentUser();



      await btn_addBookmark?.addEventListener('click', async () => {
        // this.domElement.querySelector('#btn' + item.Id + '_edit').addEventListener('click', () => {
        //localStorage.setItem("contractId", item.Id);


        try {
          await sp.web.lists.getByTitle("Marque_Pages").items.add({
            Title: item.label,
            url: `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${item.key}`,
            user: user_current.Title,
            FolderID: item.key
          })
            .then(() => {
              alert("Ajoutee dans Marque-Pages.");
              window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${item.key}`;
            });

        }
        catch (err) {
          alert(err.message);
        }
      });

      await btn_removeBookmark?.addEventListener('click', async () => {

        try {
          var items = await sp.web.lists.getByTitle("Marque_Pages").items
            .select("ID")
            .filter(`FolderID eq '${item.key}' and user eq '${user.Title}'`)
            .get();

          if (items.length === 0) {
            console.log('Item not found in Favourites list.');
            return;
          }

          // Delete the item from the Favourites list
          await sp.web.lists.getByTitle("Marque_Pages").items.getById(items[0].ID)
            .delete()
            .then(() => {
              alert("Retiree depuis Marque-Pages.");
              window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${item.key}`;
            });

        }

        catch (e) {
          alert(e.message);
        }


      });



    }

    const handleDept = async () => {

      var btn_dept: Element = document.getElementById("ajouterDept");

      let nav_html_dept: string = '';

      let nav_html_not_dept: string = '';

      nav_html_dept = `<a title="Retirer depuis department" id='${item.id}_removeDept' role="button" style="color: gold;"><svg aria-hidden="true" focusable="false" data-prefix="fas" data-icon="square-xmark" class="svg-inline--fa fa-square-xmark fa-icon fa-2x" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512">
      <path fill="currentColor" d="M64 32C28.7 32 0 60.7 0 96V416c0 35.3 28.7 64 64 64H384c35.3 0 64-28.7 64-64V96c0-35.3-28.7-64-64-64H64zm79 143c9.4-9.4 24.6-9.4 33.9 0l47 47 47-47c9.4-9.4 24.6-9.4 33.9 0s9.4 24.6 0 33.9l-47 47 47 47c9.4 9.4 9.4 24.6 0 33.9s-24.6 9.4-33.9 0l-47-47-47 47c-9.4 9.4-24.6 9.4-33.9 0s-9.4-24.6 0-33.9l47-47-47-47c-9.4-9.4-9.4-24.6 0-33.9z"></path></svg></a> `;

      nav_html_not_dept = `<a title="Ajouter dans department" id='${item.id}_addDept' role="button" style="color: gold;"><svg aria-hidden="true" focusable="false" data-prefix="far" data-icon="square-check" class="svg-inline--fa fa-square-check fa-icon fa-2x" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512">
      <path fill="currentColor" d="M211.8 339.8C200.9 350.7 183.1 350.7 172.2 339.8L108.2 275.8C97.27 264.9 97.27 247.1 108.2 236.2C119.1 225.3 136.9 225.3 147.8 236.2L192 280.4L300.2 172.2C311.1 161.3 328.9 161.3 339.8 172.2C350.7 183.1 350.7 200.9 339.8 211.8L211.8 339.8zM0 96C0 60.65 28.65 32 64 32H384C419.3 32 448 60.65 448 96V416C448 451.3 419.3 480 384 480H64C28.65 480 0 451.3 0 416V96zM48 96V416C48 424.8 55.16 432 64 432H384C392.8 432 400 424.8 400 416V96C400 87.16 392.8 80 384 80H64C55.16 80 48 87.16 48 96z"></path></svg></a>`;


      var items = await sp.web.lists.getByTitle("Department").items
        .select("ID")
        .filter(`FolderID eq '${item.key}'`)
        .get();

      if (items.length === 0) {
        btn_dept.innerHTML = nav_html_not_dept;
      } else {
        btn_dept.innerHTML = nav_html_dept;

      }


      const btn_addDept = document.getElementById(item.id + '_addDept');
      const btn_removeDept = document.getElementById(item.id + '_removeDept');

      await btn_addDept?.addEventListener('click', async () => {

        try {

          await sp.web.lists.getByTitle("Department").items.add({
            Title: item.label,
            url: `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${item.key}`,
            // user: user_current.Title,
            FolderID: item.key
          })
            .then(() => {
              alert('Dossier ajouté à la liste des départements.');
              window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${item.key}`;

            });

        }
        catch (e) {

          alert(e.message);
        }


      });

      await btn_removeDept?.addEventListener('click', async () => {

        try {
          var items = await sp.web.lists.getByTitle("Department").items
            .select("ID")
            .filter(`FolderID eq '${item.key}'`)
            .get();

          if (items.length === 0) {
            console.log('Folder not found in Department list.');
            return;
          }

          // Delete the item from the Favourites list
          await sp.web.lists.getByTitle("Department").items.getById(items[0].ID).delete();

          alert('Dossier supprimé de la liste des services.');
          window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${item.key}`;

        }

        catch (e) {

          alert(e.message);
        }
      });




    }

    const getChildrenById = async (id, items) => {

      const children = await sp.web.lists.getByTitle("Documents").items
        .select("ID, Title, ParentID, inheriting, FolderID")
        .filter(`ParentID eq '${id}'`)
        .get();

      let result = [];

      for (const child of children) {
        result.push(child);
        //  const subChildren = await getChildrenById(child.ID, items);
        //  ine fr changement
        const subChildren = await getChildrenById(child.FolderID, items);
        result = [...result, ...subChildren];
      }

      return result;
    }

    const getListItemPermissions = async (siteUrl, listName, itemId, username, password) => {
      try {
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

    const getSitePermissions = async (siteUrl, username, password) => {
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

            // Add member to permissions
            permissions.push({ type: 'member', id: principalId, role: roleName, title: title });
          }
        }

        console.log("All Site Permissions:", permissions);

        return { permissions };
      } catch (err) {
        console.error(err);
        return { permissions: [] };
      }
    }

    const generateTable = async (groups: any, x) => {
      {

        var permission_container: Element = document.getElementById("spListPermissions");

        // while (permission_container.firstChild) {
        //   permission_container.removeChild(permission_container.firstChild);
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

          html += `
            <tr>
            <td class="text-left" id="${element1.id}">${element1.id}</td>
  
            <td class="text-left" id="${element1.id}_personName">${element1.title}</td>
            
            <td class="text-center" id="${element1.id}_permission_value"> ${element1.role} </td>
  
            <td class="text-center">
            <a id="btn${element1.id}_edit" class='buttoncss' role="button">
            
            <svg aria-hidden="true" focusable="false" data-prefix="far" data-icon="trash-can" class="svg-inline--fa fa-trash-can fa-icon fa-2x" role="img" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512">
                      <path fill="currentColor" d="M160 400C160 408.8 152.8 416 144 416C135.2 416 128 408.8 128 400V192C128 183.2 135.2 176 144 176C152.8 176 160 183.2 160 192V400zM240 400C240 408.8 232.8 416 224 416C215.2 416 208 408.8 208 400V192C208 183.2 215.2 176 224 176C232.8 176 240 183.2 240 192V400zM320 400C320 408.8 312.8 416 304 416C295.2 416 288 408.8 288 400V192C288 183.2 295.2 176 304 176C312.8 176 320 183.2 320 192V400zM317.5 24.94L354.2 80H424C437.3 80 448 90.75 448 104C448 117.3 437.3 128 424 128H416V432C416 476.2 380.2 512 336 512H112C67.82 512 32 476.2 32 432V128H24C10.75 128 0 117.3 0 104C0 90.75 10.75 80 24 80H93.82L130.5 24.94C140.9 9.357 158.4 0 177.1 0H270.9C289.6 0 307.1 9.358 317.5 24.94H317.5zM151.5 80H296.5L277.5 51.56C276 49.34 273.5 48 270.9 48H177.1C174.5 48 171.1 49.34 170.5 51.56L151.5 80zM80 432C80 449.7 94.33 464 112 464H336C353.7 464 368 449.7 368 432V128H80V432z">
                      </path></svg>
            
            </a>
          </td>
            
            </tr>
            `;
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
            // alert("Remove permission with people id" + data[0]);

            try {

              var x = await getChildrenById(item.key, []);

              await sp.web.lists.getByTitle("AccessRights").items.add({
                Title: item.label.toString(),
                groupName: $("#users_name").val(),
                permission: "NONE",
                FolderID: item.id.toString(),
                PrincipleID: data[0]
              })
                .then(async () => {
                  await Promise.all(x.map(async (item_group) => {
                    await sp.web.lists.getByTitle("AccessRights").items.add({
                      Title: item_group.Title.toString(),
                      groupName: $("#users_name").val(),
                      permission: "NONE",
                      FolderID: item_group.ID,
                      PrincipleID: data[0]
                    });
                  }));

                })
                .then(async () => {
                  alert("Autorisation supprimée avec succès.");
                  await sp.web.lists.getByTitle("Documents").items.getById(item.id).update({
                    inheriting: "NO"
                  }).then(result => {
                    console.log("Item updated successfully");
                  }).catch(error => {
                    console.log("Error updating item: ", error);
                  });

                });
            }
            catch (e) {
              console.log(e.message);
            }

          });

        } else {
        }

      }
    }

    const getBasePermTest2 = async (listId, docId) => {
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

    // Function to load the next batch of pages manually

    return (
      <span

        onClick={async (event: React.MouseEvent<HTMLInputElement>) => {


          const divElement = event.currentTarget;

          if (!divElement.classList.contains('disabled')) {
            divElement.classList.add('disabled');


            var principleOfGroupAD = null;

            $("#group_name").bind('input', () => {
              const shownVal = (document.getElementById("group_name") as HTMLInputElement).value;
              // var shownVal = document.getElementById("name").value;

              const value2send = (document.querySelector<HTMLSelectElement>(`#groups option[value='${shownVal}']`) as HTMLSelectElement).dataset.value;
              principleOfGroupAD = value2send;
              console.log(value2send);
              //  $("#created_by").val(value2send);
            });


            const sidebarMenu = document.querySelector("#sidebarMenu")?.parentNode as HTMLElement | null;

            if (sidebarMenu) {
              sidebarMenu.classList.toggle("sidebar-toggle");
            }
            // const sidebarMenu = document.querySelector("div:has(>#sidebarMenu)");
            // sidebarMenu.classList.toggle("sidebar-toggle");

            const checkBoxes = document.querySelectorAll(".noCheckBox_91515d42");

            checkBoxes.forEach(box => {
              const child = box.querySelector("span");

              child.addEventListener("click", (e) => {

                const checked = document.querySelector(".checked_91515d42");

                if (checked) {
                  checked.classList.remove("checked_91515d42");
                }

                box.classList.add("checked_91515d42");
              })
            });

            var permission_container: Element = document.getElementById("spListPermissions");

            while (permission_container.firstChild) {
              permission_container.removeChild(permission_container.firstChild);
            }

            item.selectable = true;

            const checkbox_fili = document.querySelector('input[name="checkFiligrane"]') as HTMLInputElement;
            checkbox_fili.checked = true;

            const checkbox_Imprimab = document.querySelector('input[name="checkImprimab"]') as HTMLInputElement;
            checkbox_Imprimab.checked = true;


            try {
              // Create the loader element
              const loader = document.createElement("div");
              loader.id = "loader2";
              loader.innerHTML = "<div id='loader-spinner'></div>";
              document.body.appendChild(loader);

              // Execute your function here
              await fetchDocuments(Number(item.key));
              displayMetadata(item.label);
              await deleteDossier(item.id, item.label, item.parentID);
              await handleBookmark();
              await handleDept();

              

              await getBasePermTest2('df095fdf-9978-44f8-941c-23b6c095751a', item.id)
                .then(async result => {
                  // Handle the result
                  console.log('High Value:', result.high);
                  console.log('Low Value:', result.low);

                  const high = result.high;
                  const low = result.low;

                  if ((high == 2147483647 && low == 4294967295) || (high == 2147483647 && low == 4294705151)) { //full control
                    console.log("You have full control!");
                    $("#ajouterDept, #accesFolder, #bouton_delete, #editFolder, #addFolder, #ajouteDoc, #bouton_bookmark").css("display", "block");

                    const { permissions } = await getListItemPermissions('https://ncaircalin.sharepoint.com/sites/MyGed', "Documents", item.id, "mgolapkhan.ext@aircalin.nc", "musharaf2897");

                    await generateTable(permissions, Number(x));
                    console.log("PERMISSIONS ON ITEM", permissions);

                    var items = await sp.web.lists.getByTitle("Documents").items
                      .select("ID, Title, ParentID, inheriting")
                      .filter(`FolderID eq '${item.key}' and IsFolder eq 'TRUE'`)
                      .get();

                    if (items[0].inheriting === "YES") {
                      $("#inheritparagraph").css("display", "block");
                    }
                    else {
                      $("#inheritparagraph").css("display", "none");
                    }

                  }
                  else if (high == 432 && low == 1011030767) { //edit
                    console.log("You can edit!");
                    $("#ajouterDept, #accesFolder, #bouton_delete, #editFolder, #addFolder").css("display", "none");
                    $("#ajouteDoc, #bouton_bookmark").css("display", "block");
                  }
                  else if (high == 176 && low == 138612833) { //read
                    console.log("You can only read!");
                    $("#ajouterDept, #accesFolder, #bouton_delete, #editFolder, #addFolder, #ajouteDoc, #bouton_bookmark").css("display", "none");
                  }

                  else {

                  }
                })
                .catch(error => {
                  // Handle any errors
                  console.error('Error:', error);
                });

              //  const { permissions } = await getListItemPermissions('https://ncaircalin.sharepoint.com/sites/MyGed', "Documents", item.id, "mgolapkhan.ext@aircalin.nc", "musharaf2897");





              // try {
              //   const isAdmin = await checkIfUserIsAdmin(this.graphClient);
              //   const isRefUser = isAdmin || await checkIfUserIsRefUser(this.graphClient);
              //   const isGuestUser = isAdmin || await checkIfUserIsGuestUser(this.graphClient);

              //   if (isAdmin) {
              //     console.log('User is an administrator.');



              //     //  const { permissions } = await getSitePermissions('https://ncaircalin.sharepoint.com/sites/MyGed', "mgolapkhan.ext@aircalin.nc", "musharaf2897");


              //     // console.log("All the ad groups on this site", permissions);

              //   } else if (isRefUser) {
              //     console.log('User is a MYGED_REF user.');
              //     $("#ajouterDept, #accesFolder, #bouton_delete, #editFolder, #addFolder").css("display", "none");


              //   } else if (isGuestUser) {
              //     console.log('User is a MYGED_GUEST user.');
              //     $("#nav").css("display", "none");


              //   } else {

              //     console.log('User is not an administrator or a MYGED_REF or MYGED_GUEST user.');
              //     $("#nav").css("display", "block");

              //     const { permissions } = await getListItemPermissions('https://ncaircalin.sharepoint.com/sites/MyGed', "Documents", item.id, "mgolapkhan.ext@aircalin.nc", "musharaf2897");
              //     await generateTable(permissions, Number(x));
              //     console.log("PERMISSIONS ON ITEM", permissions);

              //     var items = await sp.web.lists.getByTitle("Documents").items
              //       .select("ID, Title, ParentID, inheriting")
              //       .filter(`FolderID eq '${item.key}' and IsFolder eq 'TRUE'`)
              //       .get();

              //     if (items[0].inheriting === "YES") {
              //       $("#inheritparagraph").css("display", "block");
              //     }
              //     else {
              //       $("#inheritparagraph").css("display", "none");
              //     }

              //     // const { permissions } = await getSitePermissions('https://ncaircalin.sharepoint.com/sites/MyGed', "mgolapkhan.ext@aircalin.nc", "musharaf2897");


              //     // console.log("All the ad groups on this site", permissions);

              //   }
              // } catch (error) {
              //   console.log('An error occurred while checking user permissions:', error);
              // }




              // const { permissions } = await getListItemPermissions('https://ncaircalin.sharepoint.com/sites/MyGed', "Documents", item.id, "mgolapkhan.ext@aircalin.nc", "musharaf2897");

              // await generateTable(permissions, Number(x));
              // console.log("PERMISSIONS ON ITEM", permissions);

              // Remove the loader element once the function has finished executing
              document.body.removeChild(loader);
            } catch (error) {
              console.error(error);
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
                          <a title="Archiver" 
                          role="button" id='${item.id}_deleteFolder' style="color: rgb(13, 110, 253);">
                          <img src="https://icons.iconarchive.com/icons/fa-team/fontawesome/128/FontAwesome-Box-Archive-icon.png" width="34" height="34" style="margin-top: -16px;"></img>
                          </a>`;

              delete_dossier.innerHTML = nav_html_delete_dossier;

              const btn = document.getElementById(item.id + '_deleteFolder');

              await btn?.addEventListener('click', async () => {


                if (confirm(`Voulez-vous vraiment archiver ce dossier : ${item.label} ?`)) {

                  try {

                    const list = sp.web.lists.getByTitle("Documents");

                    const i = await list.items.getById(Number(item.id)).update({
                      ParentID: 791,
                    })
                      .then(() => {
                        alert("Dossier archivé avec succès.");
                      })
                      .then(() => {
                        window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx`;
                        // window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${folderInfo[0].ParentID}`;
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

              let update_btn_dossier: string = `<button type="button" class="btn btn-primary btn_edit_dossier" id='${item.id}_update_details' style="font-size: 1em;">Modifier</button>
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
                        window.open(`https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${item.key}`, "blank");
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
          <button type="button" class="btn btn-primary add_doc" id="${item.id}_add_doc" style="font-size: 1em;">Sauvegarder</button>
          `;

              add_doc_container.innerHTML = add_btn_document;


              const btn_add_doc = document.getElementById(item.id + '_add_doc');

              await btn_add_doc?.addEventListener('click', async () => {



                const checkbox_Fili = document.querySelector<HTMLInputElement>('input[name="checkFiligrane"]');
                const checkbox_Imprimab = document.querySelector<HTMLInputElement>('input[name="checkImprimab"]');

                const value_fili = getCheckboxValue(checkbox_Fili);
                const value_impri = getCheckboxValue(checkbox_Imprimab);

                let user_current = await sp.web.currentUser();

                console.log("CURRENT USER", user_current);


                if ($("#input_revision_add").val() == '') {
                  alert("Veuillez mettre une révision avant de continuer.");
                }

                else {
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
                          createdDate: new Date().toLocaleString(),
                          IsFiligrane: value_fili,
                          IsDownloadable: value_impri
                        })
                          .then(async (iar) => {

                            const list = sp.web.lists.getByTitle("Documents");

                            await list.items.getById(iar.data.ID).attachmentFiles.add(fileName, content)

                              .then(async () => {

                                await list.items.getById(iar.data.ID).update({
                                  FolderID: parseInt(iar.data.ID),
                                  filename: fileName
                                });

                                try {

                                  var value2 = "TRUE";
                                  const folderInfo = await sp.web.lists.getByTitle('Documents').items
                                    .select("ID,ParentID,FolderID,Title,revision,IsFolder,description,attachmentUrl,IsFiligrane,IsDownloadable, inheriting")
                                    .top(5000)
                                    .filter(`FolderID eq '${item.key}' and IsFolder eq '${value2}'`)
                                    .getAll();

                                  await sp.web.lists.getByTitle("Audit").items.add({
                                    Title: iar.data.Title.toString(),
                                    DateCreated: moment().format("MM/DD/YYYY HH:mm:ss"),
                                    Action: "Creation",
                                    FolderID: iar.data.ID.toString(),
                                    Person: user_current.Title.toString()
                                  });

                                  await sp.web.lists.getByTitle("InheritParentPermission").items.add({
                                    Title: iar.data.Title.toString(),
                                    FolderID: iar.data.ID,
                                    IsDone: "NO",
                                    ParentID: Number(folderInfo[0].ID)
                                  });
                                }

                                catch (e) {
                                  alert("Erreur: " + e.message);
                                }

                              })


                            var item_id = iar.data.ID,
                              item_title = iar.data.Title;

                            return { item_id, item_title };

                          })
                          .then(({ item_id, item_title }) => {
                            alert("Document creer avec succès");
                            window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/Document.aspx?document=${item_title}&documentId=${item_id}`;
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
              });

            }

            //bouton add subfolder
            {
              var add_subfolder_container: Element = document.getElementById("add_btn_subFolder");

              let add_btn_subfolder: string = `
          <button type="button" class="btn btn-primary add_subfolder mb-2" id="${item.id}_add_btn_subfolder" style="float: right; font-size: 1em;">Ajouter sous-dossier</button>
          `;

              add_subfolder_container.innerHTML = add_btn_subfolder;


              const btn_add_subfolder = document.getElementById(item.id + '_add_btn_subfolder');


              await btn_add_subfolder?.addEventListener('click', async () => {
                var subId = null;

                if ($("#folder_name").val() == '') {
                  alert("Veuillez mettre une révision avant de continuer.")
                }

                else {
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
                          .then(async () => {

                            var value2 = "TRUE";

                            const folderInfo = await sp.web.lists.getByTitle('Documents').items
                              .select("ID,ParentID,FolderID,Title,revision,IsFolder,description,attachmentUrl,IsFiligrane,IsDownloadable, inheriting")
                              .top(5000)
                              .filter(`FolderID eq '${item.key}' and IsFolder eq '${value2}'`)
                              .getAll();

                            await sp.web.lists.getByTitle("InheritParentPermission").items.add({
                              Title: folderInfo[0].Title,
                              FolderID: iar.data.ID,
                              IsDone: "NO",
                              ParentID: Number(folderInfo[0].ID)
                            });

                            alert(`Dossier ajouté avec succès`);
                          })
                          .then(() => {
                            if (item.key !== 1) {
                              window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${item.key}`;
                            }

                            else {
                              window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx`;
                            }
                          });

                      });

                  }
                  catch (err) {
                    console.log("Erreur:", err.message);
                  }
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
          <button type="button" class="btn btn-primary add_group mb-2" style="font-size: 1em;" id=${item.id}_add_user>Ajouter</button>
          `;

              add_user_permission_container.innerHTML = add_btn_user_permission;

              const btn_add_user = document.getElementById(item.id + '_add_user');

              var peopleID = null;


              await btn_add_user?.addEventListener('click', async () => {


                var selected_permission = $("#permissions_user option:selected").val();

                var permission = 0;

                if ($("#users_name").val() === "") {
                  alert("Please select a user.");
                }
                else {

                  if (selected_permission === "ALL") {

                    permission = 1073741829;
                  }

                  else if (selected_permission === "READ") {
                    permission = 1073741826;

                  }
                  else if (selected_permission === "READ_WRITE") {
                    permission = 1073741830;

                  }


                  const user: any = await sp.web.siteUsers.getByEmail($("#users_name").val().toString())();

                  users_Permission = user;

                  console.log("USERS FOR PERMISSION", users_Permission);

                  var x = await getChildrenById(item.key, []);


                  try {
                    console.log("KEY", item.key);

                    await sp.web.lists.getByTitle("AccessRights").items.add({
                      Title: item.label.toString(),
                      groupName: $("#users_name").val(),
                      permission: $("#permissions_user option:selected").val(),
                      FolderID: item.id.toString(),
                      PrincipleID: user.Id,
                      RoleDefID: permission
                    })
                      .then(async () => {


                        await sp.web.lists.getByTitle("Documents").items.getById(item.id).update({
                          inheriting: "NO"
                        }).then(result => {
                          console.log("Item updated successfully");
                        }).catch(error => {
                          console.log("Error updating item: ", error);
                        });

                        await Promise.all(x.map(async (item) => {

                          if (item.inheriting !== "NO") {
                            await sp.web.lists.getByTitle("AccessRights").items.add({
                              Title: item.Title.toString(),
                              groupName: $("#users_name").val(),
                              permission: $("#permissions_user option:selected").val(),
                              FolderID: item.ID.toString(),
                              PrincipleID: user.Id,
                              RoleDefID: permission
                            });
                          }

                        }));


                      })
                      .then(() => {
                        alert("Autorisation ajoutée à ce dossier avec succès.")
                      })
                      // .then(() => {
                      //   sp.web.lists.getByTitle("Documents").items.getById(item.id).update({
                      //     inheriting: "NO",
                      //   }).then(result => {
                      //     console.log("Item updated successfully");
                      //   }).catch(error => {
                      //     console.log("Error updating item: ", error);
                      //   });
                      // })
                      .then(() => {
                        window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${item.key}`;
                      });

                  }

                  catch (e) {
                    alert("Erreur: " + e.message);
                  }

                }

              });




              var add_group_permission_container: Element = document.getElementById("add_btn_group");

              let add_btn_group_permission: string = `
          <button type="button" class="btn btn-primary add_group mb-2" style="font-size: 1em;" id=${item.id}_add_group>Ajouter</button>
          `;

              add_group_permission_container.innerHTML = add_btn_group_permission;

              const btn_add_group = document.getElementById(item.id + '_add_group');

              await btn_add_group?.addEventListener('click', async () => {

                var selected_permission = $("#permissions_group option:selected").val();

                var permission = 0;



                if ($("#group_name").val() === "") {
                  alert("Please select a group.");
                }
                else {

                  if (selected_permission === "ALL") {

                    permission = 1073741829;
                  }

                  else if (selected_permission === "READ") {
                    permission = 1073741826;

                  }
                  else if (selected_permission === "READ_WRITE") {
                    permission = 1073741830;

                  }

                  //  const stringGroupUsers: string[] = await getAllUsersInGroup($("#group_name").val());
                  //  console.log("TESTER GROUP USERS", stringGroupUsers);

                  add_permission_group2($("#group_name").val().toString(), permission, item.key, principleOfGroupAD);

                  await sp.web.lists.getByTitle("Documents").items.getById(item.id).update({
                    inheriting: "NO",
                  }).then(result => {
                    console.log("Item updated successfully");
                  }).catch(error => {
                    console.log("Error updating item: ", error);
                  });
                }

              });

              var inherit_permission_container: Element = document.getElementById("inheritParentFolderPermission");
              let inherit_parent_permission: string = `
                <button type="button" class="btn btn-primary add_group mb-2" style="font-size: 1em;" id=${item.id}_inheritParentPermission>Hériter les droits d'accès du parent</button>
                `;

              inherit_permission_container.innerHTML = inherit_parent_permission;

              const btn_inherit_permission = document.getElementById(item.id + '_inheritParentPermission');

              await btn_inherit_permission?.addEventListener('click', async () => {


                var x = await getChildrenById(item.key, []);


                try {
                  // console.log(item_perm.title);

                  var items = await sp.web.lists.getByTitle("Documents").items
                    .select("ID")
                    .filter(`FolderID eq '${item.parentID}' and IsFolder eq 'TRUE'`)
                    .get();



                  await sp.web.lists.getByTitle("InheritParentPermission").items.add({
                    Title: items[0].Title,
                    FolderID: item.id,
                    IsDone: "NO",
                    ParentID: Number(items[0].ID)
                  })
                    .then(async () => {
                      await Promise.all(x.map(async (item_group) => {
                        await sp.web.lists.getByTitle("InheritParentPermission").items.add({
                          Title: item_group.Title,
                          FolderID: item_group.ID,
                          IsDone: "NO",
                          ParentID: Number(items[0].ID)
                        });
                      }));

                    })
                    .then(() => {
                      console.log("ADDED PARENT");
                    })
                    .then(() => {

                      sp.web.lists.getByTitle("Documents").items.getById(item.id).update({
                        inheriting: "YES",
                      }).then(result => {
                        console.log("Item updated successfully");
                      }).catch(error => {
                        console.log("Error updating item: ", error);
                      });
                    });

                  alert("Parent permissions added.");
                  window.location.href = `https://ncaircalin.sharepoint.com/sites/MyGed/SitePages/documentation.aspx?folder=${item.key}`;

                }
                catch (e) {
                  alert(e.message);
                }
              });


            }


            //close doc upload
            {
              $("#cancel_doc").click(() => {

                $("#doc_details_add").css("display", "none");
              });
            }


            divElement.classList.remove('disabled');
          }
        }

        }

      >


        {

          <FontAwesomeIcon icon={item.icon} className="fa-icon" ></FontAwesomeIcon >
        }

        &nbsp;

        {item.label}

      </span >
    );

  }

}


