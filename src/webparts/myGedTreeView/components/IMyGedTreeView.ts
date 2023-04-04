
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ITreeItem } from "@pnp/spfx-controls-react/lib/TreeView";
import { MSGraphClient, MSGraphClientFactory } from '@microsoft/sp-http';
import { IconDefinition } from "@fortawesome/fontawesome-svg-core";



export interface IMyGedTreeViewState {
  //TreeLinks: ITreeItem[];
  TreeLinks: any[];
  parentIDArray: any[];
  isLoaded: any;
  selectedKey: any;
  isToggledOn: boolean;
  isToggleOnDept: boolean;

}

export interface IMyGedTreeViewProps {
  description: string;
  context: any | null;
  msGraphClientFactory: MSGraphClientFactory,
}