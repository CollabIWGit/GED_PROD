
import { ITreeItem } from "@pnp/spfx-controls-react/lib/TreeView";


export interface IMyGedTreeViewState {
    TreeLinks: ITreeItem[];
  }
  
  export interface IMyGedTreeViewProps {
    description: string;
    context: any | null;
  }