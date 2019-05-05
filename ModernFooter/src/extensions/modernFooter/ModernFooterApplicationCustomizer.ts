import * as React from "react";
import * as ReactDom from "react-dom";
import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from "@microsoft/sp-application-base";
import { Dialog } from "@microsoft/sp-dialog";
import { IPortalFooterProps, PortalFooter } from "./components/PortalFooter";
import { ILinkGroup } from "./components/PortalFooter/ILinkGroup";
import { ILinkListItem } from "./ILinkListItem";
import pnp, { Web } from "@pnp/pnpjs";

import { autobind } from "@uifabric/utilities";
import { IPortalFooterEditResult } from "./components/PortalFooter/IPortalFooterEditResult";

import * as strings from "ModernFooterApplicationCustomizerStrings";
import {
  SPHttpClient,
  ISPHttpClientOptions,
  SPHttpClientResponse
} from "@microsoft/sp-http";

const LOG_SOURCE: string = "ModernFooterApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IModernFooterApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ModernFooterApplicationCustomizer extends BaseApplicationCustomizer<
  IModernFooterApplicationCustomizerProperties
> {
  private static _bottomPlaceholder?: PlaceholderContent;

  @override
  public onInit(): Promise<void> {
    console.log(`Hello from ModernFooter`);

    // call render method for generating the needed html elements
    return this._renderPlaceHolders();
    //return Promise.resolve();
  }

  private _handleDispose(): void {
    console.log(
      "[PortalFooterApplicationCustomizer._onDispose] Disposed custom bottom placeholder."
    );
  }

  protected getContent(): Promise<any> {
    var httpClientOptions: ISPHttpClientOptions = {};

    httpClientOptions.headers = {
      Accept: "application/json;odata=nometadata",
      "odata-version": ""
    };

    return this.context.spHttpClient
      .get(
        `https://cargillonline.sharepoint.com/sites/InnovatorsStudio/_api/web/lists('7b028aab-0cca-4d0b-9ee6-79976c9cd721')/items?$select=Id,Title,FooterLinkGroup,FooterUrl`,
        SPHttpClient.configurations.v1,
        httpClientOptions
      )
      .then(
        (response: SPHttpClientResponse): Promise<{ value: any[] }> => {
          return response.json();
        }
      );
  }

  private loadLinks(): Promise<ILinkGroup[]> {
    return new Promise<ILinkGroup[]>(resolve => {
      let web = new Web(
        "https://cargillonline.sharepoint.com/sites/InnovatorsStudio/"
      );

      web.get().then(w => {
        w.lists.get().then(l => {
          console.log(l);
        });
        // get the links from the source list
        let items: ILinkListItem[] = w.lists
          .getByTitle("MIIFooter")
          .items.select("Title", "FooterLinkGroup", "FooterUrl")
          .orderBy("FooterLinkGroup", true)
          .orderBy("Title", true)
          .get()
          .then((footerItems: any[]) => {
            console.log(footerItems);
            // prepare the result variable
            let result: ILinkGroup[] = [];

            footerItems.map((v, i, a) => {
              // in case we have a new group title
              if (
                result.length === 0 ||
                v.FooterLinkGroup !== result[result.length - 1].title
              ) {
                // create the new group and add the current item
                result.push({
                  title: v.FooterLinkGroup,
                  links: [
                    {
                      title: v.Title,
                      url: v.FooterUrl.Url
                    }
                  ]
                });
              } else {
                // or add the current item to the already existing group
                result[result.length - 1].links.push({
                  title: v.Title,
                  url: v.FooterUrl.Url
                });
              }
            });

            resolve(result);
          })
          .catch(console.log);
      });
    });
  }

  private async loadLinks2(): Promise<ILinkGroup[]> {
    const { sp } = await import(/* webpackChunkName: 'pnp-sp' */
    "@pnp/sp");

    pnp.sp.web.lists
      .getByTitle("MIIFooter")
      .items.get()
      .then((items: any[]) => {
        console.log(items);
      });

    // prepare the result variable
    let result: ILinkGroup[] = [];

    let web = new Web(
      "https://cargillonline.sharepoint.com/sites/InnovatorsStudio/"
    );

    web.get().then(w => {
      console.log(w.Url);
      console.log(w.lists);
      console.log(web.lists);

      // get the links from the source list
      let items: ILinkListItem[] = w.lists
        .getByTitle("MIIFooter")
        .items.select("Title", "FooterLinkGroup", "FooterUrl")
        .orderBy("FooterLinkGroup", true)
        //.orderBy("Title", true)
        .get()
        .catch(console.log);

      // map the list items to the results
      items.map((v, i, a) => {
        // in case we have a new group title
        if (
          result.length === 0 ||
          v.FooterLinkGroup !== result[result.length - 1].title
        ) {
          // create the new group and add the current item
          result.push({
            title: v.FooterLinkGroup,
            links: [
              {
                title: v.Title,
                url: v.FooterUrl.Url
              }
            ]
          });
        } else {
          // or add the current item to the already existing group
          result[result.length - 1].links.push({
            title: v.Title,
            url: v.FooterUrl.Url
          });
        }
      });
    });

    return result;
  }

  private async _renderPlaceHolders(): Promise<void> {
    // check if the application customizer has already been rendered
    if (!ModernFooterApplicationCustomizer._bottomPlaceholder) {
      console.log("bottomPlaceholder not rendered yet");
      // create a DOM element in the bottom placeholder for the application customizer to render
      ModernFooterApplicationCustomizer._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._handleDispose }
      );
    }

    // if the top placeholder is not available, there is no place in the UI
    // for the app customizer to render, so quit.
    // top placeholder..
    let topPlaceholder: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Top
    );
    if (topPlaceholder) {
      topPlaceholder.domElement.innerHTML = ``;
    }

    // prepare the result variable
    let result: ILinkGroup[] = [];

    this.getContent().then(stuff => {
      console.log("returned getcontent");
      //console.log(stuff);
      console.log(stuff.value);

      stuff.value.map((v, i, a) => {
        // in case we have a new group title
        if (
          result.length === 0 ||
          v.FooterLinkGroup !== result[result.length - 1].title
        ) {
          // create the new group and add the current item
          result.push({
            title: v.FooterLinkGroup,
            links: [
              {
                title: v.Title,
                url: v.FooterUrl.Url
              }
            ]
          });
        } else {
          // or add the current item to the already existing group
          result[result.length - 1].links.push({
            title: v.Title,
            url: v.FooterUrl.Url
          });
        }
      });

      console.log(result);

      const element: React.ReactElement<
        IPortalFooterProps
      > = React.createElement(PortalFooter, {
        links: result
      });

      // render the UI using a React component
      ReactDom.render(
        element,
        ModernFooterApplicationCustomizer._bottomPlaceholder.domElement
      );
    });

    //const links: ILinkGroup[] = await this.loadLinks();
    //console.log(links);
  }
}
