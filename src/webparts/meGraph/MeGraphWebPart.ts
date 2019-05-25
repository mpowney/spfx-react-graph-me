import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownProps,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { MSGraphClient } from '@microsoft/sp-http';
import { Options as GraphClientOptions } from '@microsoft/microsoft-graph-client';
import * as Msal from 'msal';

import * as strings from 'MeGraphWebPartStrings';
import MeGraph from './components/MeGraph';
import { IMeGraphProps } from './components/IMeGraphProps';

export interface IMeGraphWebPartProps {
  graphEndpoint: string;
}

export default class MeGraphWebPart extends BaseClientSideWebPart<IMeGraphWebPartProps> {

  private _token: string;

  public async onInit(): Promise<void> {

    const scopes = ["People.Read.All", "User.Read.All"];
    const msalConfig = {
      auth: {
          clientId: 'cc011966-b43c-4b92-b820-0d00204e8e21',
          authority: `https://login.microsoftonline.com/b0327ae0-4c3a-4405-8847-f8c132aa5bb0`
      },
      cache: {
        cacheLocation: Msal.Constants.cacheLocationLocal
      },
    };

    const tokenRequest = {
      scopes: scopes
    };

    const msalInstance = new Msal.UserAgentApplication(msalConfig);

      if (msalInstance.getAccount()) {

        try {
          this._token = (await msalInstance.acquireTokenSilent(tokenRequest)).accessToken;
          console.log(`acquireTokenSilent accessToken ${this._token}`);
        }
        catch (err) {
          try {
            this._token = (await msalInstance.acquireTokenPopup(tokenRequest)).accessToken;
            console.log(`acquireTokenPopup accessToken ${this._token}`);
          }
          catch (err) {
            console.error(err);
          }
        }

      } else {

        try {
          const idToken = (await msalInstance.loginPopup(tokenRequest)).idToken;
          console.log(`loginPopup idToken:`);
          console.log(idToken);
          this._token = (await msalInstance.acquireTokenPopup(tokenRequest)).accessToken;
          console.log(`acquireTokenPopup accessToken ${this._token}`);
      }
        catch (err) {
          console.error(err);
        }

      }

    }

  public render(): void {

    this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        
          const graphClientConfig: GraphClientOptions = {
            authProvider: (done) => {
              done(undefined, this._token)
            }
          };

          client
            .api(`/me${this.properties.graphEndpoint}`, graphClientConfig)
            .get((error, response: any, rawResponse?: any) => {

              const element: React.ReactElement<IMeGraphProps > = React.createElement(
                MeGraph,
                {
                  selectedEndpoint: this.properties.graphEndpoint,
                  graphData: response,
                  isLoading: false,
                  graphClient: client
                }
              );
          
              ReactDom.render(element, this.domElement);
          });

        });


    const placeholderElement: React.ReactElement<IMeGraphProps > = React.createElement(
      MeGraph,
      {
        selectedEndpoint: this.properties.graphEndpoint,
        graphData: null,
        isLoading: true,
        graphClient: null
      }
    );

    ReactDom.render(placeholderElement, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    const graphEndpointOptions: IPropertyPaneDropdownOption[] = [
      { key: `/people`, text: `People` }
    ];

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
                PropertyPaneDropdown('graphEndpoint', {
                  label: strings.GraphEndpointDropdownLabel,
                  options: graphEndpointOptions
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
