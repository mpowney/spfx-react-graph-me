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

import * as strings from 'MeGraphWebPartStrings';
import MeGraph from './components/MeGraph';
import { IMeGraphProps } from './components/IMeGraphProps';
import { MSGraphClient } from '@microsoft/sp-http';
import { SPComponentLoader } from "@microsoft/sp-loader";

const COMPONENT_ID: any = {
  SP_WEBPART_SHARED: "914330ee-2df2-4f6e-a858-30c23a812408",
  SP_RTE: "8404d628-4817-4b3a-883e-1c5a4d07892e",
  SP_SUITE_NAV: "f8a8ad94-4cf3-4a19-a76b-1cec9da00219",
  SP_COMPONENT_UTILITIES: "8494e7d7-6b99-47b2-a741-59873e42f16f",
  ODSP_UTILITIES_BUNDLE: "cc2cc925-b5be-41bb-880a-f0f8030c6aff",
  SP_PAGE_CONTEXT: "1c4541f7-5c31-41aa-9fa8-fbc9dc14c0a8"
}

export interface IMeGraphWebPartProps {
  graphEndpoint: string;
}

export default class MeGraphWebPart extends BaseClientSideWebPart<IMeGraphWebPartProps> {

  public render(): void {

    this.loadSPComponentById(COMPONENT_ID.SP_WEBPART_SHARED).then((library: any) => {
      console.log('SP_WEBPART_SHARED:');
      console.log(library);
    });

    this.loadSPComponentById(COMPONENT_ID.SP_RTE).then((library: any) => {
      console.log('SP_RTE:');
      console.log(library);
    });

    this.loadSPComponentById(COMPONENT_ID.SP_SUITE_NAV).then((library: any) => {
        console.log('SP_SUITE_NAV:');
        console.log(library);
        console.log(`SuiteNavManagerConfiguration.isSearchBoxInHeaderFlighted`);
        console.log(library.SuiteNavManagerConfiguration.isSearchBoxInHeaderFlighted());
        console.log(`SuiteNavManager().loadSuiteNav.get()`);
        console.log(library.SuiteNavManager().loadSuiteNav().get());
    });

    this.loadSPComponentById(COMPONENT_ID.SP_COMPONENT_UTILITIES).then((library: any) => {
        console.log('SP_COMPONENT_UTILITIES:');
        console.log(library);
        console.log('SPUtility.getUserPhotoUrl():');
        console.log(library.SPUtility.getUserPhotoUrl(this.context.pageContext.user.loginName))
    });

    this.loadSPComponentById(COMPONENT_ID.ODSP_UTILITIES_BUNDLE).then((library: any) => {
        console.log('ODSP_UTILITIES_BUNDLE:');
        console.log(library);
    });

    this.loadSPComponentById(COMPONENT_ID.SP_PAGE_CONTEXT).then((library: any) => {
        console.log('SP_PAGE_CONTEXT:');
        console.log(library);
    });


    this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {

        client
          .api(`/me${this.properties.graphEndpoint}`)
          .get((error, response: any, rawResponse?: any) => {

            const element: React.ReactElement<IMeGraphProps > = React.createElement(
              MeGraph,
              {
                selectedEndpoint: this.properties.graphEndpoint,
                graphData: response,
                isLoading: false,
                graphClient: client,
                serviceScope: this.context.serviceScope
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
        graphClient: null,
        serviceScope: this.context.serviceScope
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

    /**
     * Load SPFx component by id, SPComponentLoader is used to load the SPFx components
     * @param componentId - componentId, guid of the component library
     */
    private loadSPComponentById(componentId: string) {

      const w = (window as any);
      w.cachedspComponent = w.cachedspComponent || {};
      w.cachedspComponent[componentId] = w.cachedspComponent[componentId] || new Promise((resolve, reject) => {
          SPComponentLoader.loadComponentById(componentId).then((component: any) => {
              resolve(component);
          }).catch((error) => {
              console.error(`MeGraphWebPart.tsx`, error, this.context.serviceScope);
              resolve(null);
          });
      });

      return w.cachedspComponent[componentId];
  }

}
