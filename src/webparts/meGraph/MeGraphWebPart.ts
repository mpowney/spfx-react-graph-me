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
import { Person } from '@microsoft/microsoft-graph-types';

export interface IMeGraphWebPartProps {
  graphEndpoint: string;
}

export default class MeGraphWebPart extends BaseClientSideWebPart<IMeGraphWebPartProps> {

  public render(): void {

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
        graphData: MeGraphWebPart.getLoadingMockData(this.properties),
        isLoading: true,
        graphClient: null
      }
    );

    ReactDom.render(placeholderElement, this.domElement);
  }

  private static getLoadingMockData(props: IMeGraphWebPartProps) {
    switch (props.graphEndpoint) {
      case '/people':
        
        const mockPerson: Person = {
          givenName: 'Josh',
          surname: 'Smith',
          displayName: 'Josh Smith',
          jobTitle: 'Sales Consultant'
        }

        return {value: [mockPerson, mockPerson, mockPerson, mockPerson, mockPerson, mockPerson] };

    }
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
