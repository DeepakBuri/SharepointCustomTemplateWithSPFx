import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'SpFxWebPartStrings';
import SpFx from './components/SpFx';
import { ISpFxProps } from './components/ISpFxProps';

export interface ISpFxWebPartProps {
  description: string;
  primarySystemAccount: string;
  recordType: string;
  owner: string;
  parantroom:string;
  restrictions:string;
}

export default class SpFxWebPart extends BaseClientSideWebPart<ISpFxWebPartProps> {


  public render(): void {
    const element: React.ReactElement<ISpFxProps> = React.createElement(
      SpFx,
      {
        description: this.properties.description,
        context:this.context,
        primarySystemAccount: this.properties.primarySystemAccount,
        recordType: this.properties.recordType,
        owner: this.properties.owner,
        parantroom:this.properties.parantroom,
        restrictions:this.properties.restrictions,
      }
    );

   
    ReactDom.render(element, this.domElement);
  }


  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

 

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          // header: {
          //   description: strings.PropertyPaneDescription,
          // },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('primarySystemAccount', {
                  label: 'Primary System Account'
                }),
                PropertyPaneTextField('description', {
                  label: 'Description'
                }),
                PropertyPaneTextField('recordType', {
                  label: 'Record Type'
                }),
                PropertyPaneTextField('parantroom', {
                  label: 'Parant Room'
                }),
                PropertyPaneTextField('owner', {
                  label: 'Owner'
                }),
                PropertyPaneTextField('restrictions', {
                  label: 'Restrictions'
                })
              ]
            }
          ]
        }
      ]
    };
  }
} 