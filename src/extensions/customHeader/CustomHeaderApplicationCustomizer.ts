import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent
  // PlaceholderContent,
  // PlaceHolderContentPlaceHolder,
} from '@microsoft/sp-application-base';
import { override } from "@microsoft/decorators";

import * as strings from 'CustomHeaderApplicationCustomizerStrings';

const LOG_SOURCE: string = 'CustomHeaderApplicationCustomizer';

import styles from './loc/CustomHeader.module.scss';
import serviceAPI from '../../APIs/APISCalls';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICustomHeaderApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CustomHeaderApplicationCustomizer
  extends BaseApplicationCustomizer<ICustomHeaderApplicationCustomizerProperties> {
    @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Add a placeholder to the header
    this.context.placeholderProvider.changedEvent.add(this, this.renderHeader);

    return Promise.resolve();
  }


  private renderHeader(): void {
    // Get the header placeholder
    const header: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(
      1,
      { onDispose: this.onDispose }
    );
    

    if (header.domElement) {
      //if dom element is present with same id then remove it and add new one
      //find the element by id
      const existingElement = document.getElementById('spTopPlaceholder');
      //replace the innerHTML of the element with new content
      if(existingElement){
        existingElement.innerHTML = `
        <div class="${styles.app}">
          <div class="${styles.top}">
              MyGlobalFoundries - Customer Team Room: This is a collaboration space for
            GlobalFoundries and customer teams.
          </div>
        </div>`;
      }
      else{
        header.domElement.innerHTML = `
        <div class="${styles.app}">
          <div class="${styles.top}">
              MyGlobalFoundries - Customer Team Room: This is a collaboration space for
            GlobalFoundries and customer teams.
          </div>
        </div>`;
      }

    
  }
  

    const SiteHeaderTitle= document.querySelector('.ms-TooltipHost');
    // append the sapn to the SiteHeaderTitle
    if (SiteHeaderTitle) {
      const newTitle = document.createElement('span');
      newTitle.tabIndex = 0;
      newTitle.className = styles.brandtext2;
      newTitle.innerHTML = "Customer Team Room" + " | " + this.context.pageContext.web.title;
      SiteHeaderTitle.innerHTML = '';
      SiteHeaderTitle.appendChild(newTitle);
    }
    serviceAPI.addPageToSitePages(
      this.context.pageContext.web.absoluteUrl,
      "Root.aspx",
    ).then((pageItem) => {
      console.log(pageItem);
    });
    
}
}
