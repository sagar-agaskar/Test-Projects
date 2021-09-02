import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpfxNoJsWebPart.module.scss';
import * as strings from 'SpfxNoJsWebPartStrings';
import { SPHttpClient } from '@microsoft/sp-http';
import { SPHttpClientResponse } from '@microsoft/sp-http';
import { Environment, EnvironmentType} from '@microsoft/sp-core-library';

export interface ISpfxNoJsWebPartProps {
  description: string;
}

export interface ISPLists{
  value : ISPList[];
}

export interface ISPList{

  gender : string,
  address: string,
  contactNumber : number
}

export default class SpfxNoJsWebPart extends BaseClientSideWebPart<ISpfxNoJsWebPartProps> {

  private _getListData : Promise<ISPLists>
  {

    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('Customer')/Items",SPHttpClient.configurations.v1)

        .then((response: SPHttpClientResponse) => 

        {

        return response.json();

        });

  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.spfxNoJs }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
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


