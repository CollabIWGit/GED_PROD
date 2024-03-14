
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

import { sp, List, IItemAddResult } from "@pnp/sp/presets/all";

var myVar;
var SP;



export interface IMyGedTreeViewWebPartProps {
  description: string;
}


export default class MyGedTreeViewWebPart extends BaseClientSideWebPart<IMyGedTreeViewWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = ''; s

  protected onInit(): Promise<void> {

    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context,
        globalCacheDisable: true
      });
    });
  }



  // public render(): void {

  //   const element: React.ReactElement<IMyGedTreeViewProps> = React.createElement(

  //     MyGedTreeView,
  //     {
  //       description: this.properties.description,
  //       context: this.context,
  //       msGraphClientFactory: this.context.msGraphClientFactory
  //     },
  //   );


  //   ReactDom.render(element, this.domElement);
  //   this.require_libraries();

  //   SPComponentLoader.loadScript('//code.jquery.com/jquery-3.3.1.slim.min.js', {
  //     globalExportsName: 'jQuery'
  //   }).then(() => {
  //     return SPComponentLoader.loadScript('//cdn.jsdelivr.net/npm/popper.js@1.14.7/dist/umd/popper.min.js');
  //   }).then(() => {
  //     return SPComponentLoader.loadScript('//cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.1.0/js/bootstrap.min.js');
  //   });


  // }

  public render(): void {

    if (!this.renderedOnce) {

      const element = React.createElement(MyGedTreeView, {
        description: this.properties.description,
        context: this.context,
        msGraphClientFactory: this.context.msGraphClientFactory
      });

      // Promise.all([
      //   this.require_libraries(),
      //   // SPComponentLoader.loadScript('//code.jquery.com/jquery-3.3.1.slim.min.js', { globalExportsName: 'jQuery' }),
      //   // SPComponentLoader.loadScript('//cdn.jsdelivr.net/npm/popper.js@1.14.7/dist/umd/popper.min.js'),
      //   // SPComponentLoader.loadScript('//cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.1.0/js/bootstrap.min.js'),

      // ]).then(() => {
      //   ReactDom.render(element, this.domElement);
      // });

      Promise.all([
        this.require_libraries(),
        // SPComponentLoader.loadScript('//code.jquery.com/jquery-3.3.1.slim.min.js', { globalExportsName: 'jQuery' }),
        // SPComponentLoader.loadScript('//cdn.jsdelivr.net/npm/popper.js@1.14.7/dist/umd/popper.min.js'),
        // SPComponentLoader.loadScript('//cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.1.0/js/bootstrap.min.js'),

      ]).then(() => {
        ReactDom.render(element, this.domElement);
      });

      // Promise.all([
      //   SPComponentLoader.loadScript('https://ajax.googleapis.com/ajax/libs/jquery/3.6.3/jquery.min.js'), // jQuery
      //   SPComponentLoader.loadScript('//cdn.jsdelivr.net/npm/popper.js@1.14.7/dist/umd/popper.min.js'), // Popper.js
      //   SPComponentLoader.loadScript('//cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.4.1/js/bootstrap.min.js'), // Bootstrap
      //   SPComponentLoader.loadScript("https://code.jquery.com/ui/1.12.1/jquery-ui.js"),
      //   this.require_libraries(), // jQuery UI
      //   // SPComponentLoader.loadScript('//cdn.datatables.net/1.10.25/js/jquery.dataTables.min.js'), // DataTables
      // ]).then(function() {
      //   console.log("Scripts loaded successfully");
      //   // ReactDom.render(element, this.domElement);
      //   // All scripts have been loaded successfully
      // }).catch(function(error) {
      //   // An error occurred while loading the scripts
      //   console.error("Error loading scripts: "+error);
      // });

      // ReactDom.render(element, this.domElement);
    }
  }

  // private require_libraries() {
  //   // require('./../../common/js/jquery.min');
  //   // require('./../../common/js/popper');
  //   // require('./../../common/js/bootstrap.min');
  //   // require('./../../common/js/main');
  //   // require('./../../common/css/common.css');
  //   require('./../../common/css/bugfix.css');
  //   require('./../../common/css/bugfix2.css');

  //   require('./../../common/css/sidebar.css');
  //   require('./../../common/css/pagecontent.css');
  //   require('./../../common/css/spinner.css');
  //   require('./../../common/css/responsive.css');
  //   require('./../../common/css/forms.css');
  // }


  private require_libraries() {
    require('./../../common/js/jquery.min');
    require('./../../common/js/popper');
    require('./../../common/js/bootstrap.min');
    require('./../../common/js/main');
    // require('./../../common/css/common.css');
    require('./../../common/css/bugfix.css');
    require('./../../common/css/bugfix2.css');

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
