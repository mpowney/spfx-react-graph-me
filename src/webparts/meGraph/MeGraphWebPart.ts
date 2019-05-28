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
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from "@microsoft/sp-component-base";

export interface IMeGraphWebPartProps {
  graphEndpoint: string;
  title: string;
}

export default class MeGraphWebPart extends BaseClientSideWebPart<IMeGraphWebPartProps> {

  private _themeProvider: ThemeProvider | null = null;
  private _themeVariant: any;
  
  protected onInit(): Promise<void> {
    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    return super.onInit();
  }

  /**
   * Update the current theme variant reference and re-render.
   *
   * @param args The new theme
   */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
      this._themeVariant = args.theme;
      this.render();
  }

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
                graphClient: client,
                title: this.properties.title,
                displayMode: this.displayMode,
                updateTitleProperty: this.updateTitleProperty.bind(this),
                themeVariant: this._themeVariant
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
        title: this.properties.title,
        displayMode: this.displayMode,
        updateTitleProperty: this.updateTitleProperty.bind(this),
        themeVariant: this._themeVariant
}
    );

    ReactDom.render(placeholderElement, this.domElement);
  }

  public updateTitleProperty = (value: string): void => {
    this.properties.title = value;
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
