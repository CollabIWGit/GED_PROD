import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'MyGedTreeViewWebPartStrings';
import MyGedTreeView from './components/MyGedTreeView';
import { IMyGedTreeViewProps, IMyGedTreeViewState } from './components/IMyGedTreeView';

var myVar;



export interface IMyGedTreeViewWebPartProps {
  description: string;
}


export default class MyGedTreeViewWebPart extends BaseClientSideWebPart<IMyGedTreeViewWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';s

  protected onInit(): Promise<void> {

    return super.onInit();
  }

  public render(): void {

    
    const element: React.ReactElement<IMyGedTreeViewProps> = React.createElement(
      
      MyGedTreeView,
      {
        description: this.properties.description,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
    
  }


  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
