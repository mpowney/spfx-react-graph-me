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

import { PropertyFieldGraphPeople, IPropertyFieldGraphPeopleProps } from '../../propertyFields/graphPeople';
import * as strings from 'MeGraphWebPartStrings';
import MeGraph from './components/MeGraph';
import { IMeGraphProps } from './components/IMeGraphProps';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import IGraphPeopleSettings, { ShowUser } from '../../propertyFields/graphPeople/IGraphPeopleSettings';

export interface IMeGraphWebPartProps {
  graphEndpoint: string;
  peopleConfig: IGraphPeopleSettings;
}

export default class MeGraphWebPart extends BaseClientSideWebPart<IMeGraphWebPartProps> {

  private static getApiEndpoint(props: IMeGraphWebPartProps): string {

    switch (props.graphEndpoint) {
      case '/people':
        let api: string = props.graphEndpoint;
        if (ShowUser[props.peopleConfig.ShowUser] === ShowUser[ShowUser.SpecifiedUser]) {
          api = `/users/${props.peopleConfig.SpecifiedUsername}/people`;
        }
        else {
          api = `/me/people`;
        }
        if (props.peopleConfig.FilterOnlyUsers) {
          api += `?$filter=personType/subclass eq 'OrganizationUser'`;
        }
        return api;
        break;
      default:
        return props.graphEndpoint;
    }


  }

  public render(): void {

    this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {

        client
          .api(MeGraphWebPart.getApiEndpoint(this.properties))
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
                }),
                PropertyFieldGraphPeople('peopleConfig', {
                  label: '',
                  properties: this.properties
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
