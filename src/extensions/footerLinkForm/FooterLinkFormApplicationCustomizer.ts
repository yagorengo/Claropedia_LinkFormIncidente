import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import * as strings from 'FooterLinkFormApplicationCustomizerStrings';
import { sp } from "@pnp/sp/presets/all";
import * as React from 'react';
import * as ReactDOM from "react-dom";  
import ReactFooter, { IReactFooterProps } from "./ReactFooter";  
const LOG_SOURCE: string = 'FooterLinkFormApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IFooterLinkFormApplicationCustomizerProperties {
  // This is an example; replace with your own property
  Bottom: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class FooterLinkFormApplicationCustomizer
  extends BaseApplicationCustomizer<IFooterLinkFormApplicationCustomizerProperties> {
    private static footerPlaceholder: PlaceholderContent;
    private _bottomPlaceholder: PlaceholderContent | undefined;
    private pathName: string = '';

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    sp.setup(this.context);
    this.pathName = this.context.pageContext.site.serverRequestPath;
    
    this.context.application.navigatedEvent.add(this, () => {
      this.loadReactComponent();
    });
    this.render();

    return Promise.resolve();
  }

  
  private async loadReactComponent() {
    
    if (FooterLinkFormApplicationCustomizer.footerPlaceholder && FooterLinkFormApplicationCustomizer.footerPlaceholder.domElement) {
      const element: React.ReactElement<IReactFooterProps> = React.createElement(ReactFooter, {
        context: this.context
      });

      ReactDOM.render(element, FooterLinkFormApplicationCustomizer.footerPlaceholder.domElement);
    }
    else {
      console.log('DOM element of the header is undefined. Start to re-render.');
      this.render();
    }
  }

  private render(): void {
    if (this.context.placeholderProvider.placeholderNames.indexOf(PlaceholderName.Bottom) !== -1) {
      if (!FooterLinkFormApplicationCustomizer.footerPlaceholder || !FooterLinkFormApplicationCustomizer.footerPlaceholder.domElement) {
        FooterLinkFormApplicationCustomizer.footerPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom, {
          onDispose: this._onDispose
        });
      }

      this.loadReactComponent();
    }
    else {
      console.log(`The following placeholder names are available`, this.context.placeholderProvider.placeholderNames);
    }
  }

  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
  
 
}
