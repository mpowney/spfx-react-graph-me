import { MSGraphClient } from '@microsoft/sp-http';
import { DisplayMode } from '@microsoft/sp-core-library';
import { IReadonlyTheme } from "@microsoft/sp-component-base";

export interface IMeGraphProps {
  selectedEndpoint: string;
  graphData: any;
  isLoading: boolean;
  graphClient: MSGraphClient;
  displayMode: DisplayMode;
  title: string;
  updateTitleProperty: (value: string) => void;
  themeVariant: IReadonlyTheme;
}
