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
import 'datatables.net';
import * as moment from 'moment';
import 'downloadjs';
import { SPComponentLoader } from '@microsoft/sp-loader';


var myVar;




export interface IMyGedTreeViewWebPartProps {
  description: string;
}


export default class MyGedTreeViewWebPart extends BaseClientSideWebPart<IMyGedTreeViewWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = ''; s

  protected onInit(): Promise<void> {

    return super.onInit();
  }



  public render(): void {



    const element: React.ReactElement<IMyGedTreeViewProps> = React.createElement(

      MyGedTreeView,
      {
        description: this.properties.description,
        context: this.context,
        msGraphClientFactory: this.context.msGraphClientFactory
      },


    );


    ReactDom.render(element, this.domElement);
    this.require_libraries();

    SPComponentLoader.loadScript('https://code.jquery.com/jquery-3.3.1.slim.min.js', {
      globalExportsName: 'jQuery'
    }).then(() => {
      return SPComponentLoader.loadScript('https://cdn.jsdelivr.net/npm/popper.js@1.14.7/dist/umd/popper.min.js');
    }).then(() => {
      return SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.1.0/js/bootstrap.min.js');
    });


  }



  private require_libraries() {
    //SideMenuUtils.buildSideMenu(this.context.pageContext.web.absoluteUrl);
    require('./../../common/js/jquery.min');
    require('./../../common/js/popper');
    require('./../../common/js/bootstrap.min');
    require('./../../common/js/main');


    require('./../../common/css/common.css');
    // require('./../../common/css/minBootstrap.css');
    require('./../../common/css/sidebar.css');
    require('./../../common/css/pagecontent.css');
    require('./../../common/css/spinner.css');
    require('./../../common/css/responsive.css');
    require('./../../common/css/forms.css');
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
