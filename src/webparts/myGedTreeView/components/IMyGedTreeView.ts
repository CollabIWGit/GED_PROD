
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ITreeItem } from "@pnp/spfx-controls-react/lib/TreeView";
import { MSGraphClient, MSGraphClientFactory } from '@microsoft/sp-http';



export interface IMyGedTreeViewState {
  //TreeLinks: ITreeItem[];
  TreeLinks: any[];

  // TreeLinks: any;
  // data: any | null; //ine zouter
}

export interface IMyGedTreeViewProps {
  description: string;
  context: any | null;
  msGraphClientFactory: MSGraphClientFactory,
}