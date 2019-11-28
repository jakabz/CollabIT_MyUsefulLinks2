import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import * as strings from 'MyUsefulLinksWebPartStrings';
import MyUsefulLinks from './components/MyUsefulLinks';
import { IMyUsefulLinksProps } from './components/IMyUsefulLinksProps';

export interface IMyUsefulLinksWebPartProps {
  title: string;
  absoluteUrl: string;
  myLinks: any;
  render: any;
  tsThis: any;
}

export default class MyUsefulLinksWebPart extends BaseClientSideWebPart<IMyUsefulLinksWebPartProps> {
  
  private listResult:any;
  private listInit:boolean = false;
  private defaultListResult:any;
  private deaultListInit:boolean = false;

  public render(): void {
    
    if(!this.listInit){
      this._startData(this);
    }

    if(!this.deaultListInit){
      this._startDefaultData(this);
    }
    
    const element: React.ReactElement<IMyUsefulLinksProps > = React.createElement(
      MyUsefulLinks,
      {
        title: this.properties.title,
        absoluteUrl: this.context.pageContext.site.absoluteUrl,
        defaultLinks: this.defaultListResult,
        myLinks: this.listResult,
        render: this._startData,
        tsThis: this
      }
    );
    if(this.listInit && this.deaultListInit){
      ReactDom.render(element, this.domElement);
    }
  }

  public _startData(self): void {
      self._getListData('My Useful Links').then((response) => {
        self.listResult = response.value;
        self.listInit = true;
        self.render();
      });
  }

  public _startDefaultData(self): void {
    self._getListData('Useful Links').then((response) => {
      self.defaultListResult = response.value;
      self.deaultListInit = true;
      self.render();
    });
}

  public _getListData(listName:string): Promise<any> {
    let query = '';
    query += '$filter=AuthorId eq ' + this.context.pageContext.legacyPageContext.userId +'&';
    query += '$orderby=Position asc';
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/Lists/GetByTitle('`+listName+`')/Items?` + query, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
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
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
