import { MSGraphClient } from '@microsoft/sp-http';

export interface IMeGraphProps {
  selectedEndpoint: string;
  graphData: any;
  isLoading: boolean;
  graphClient: MSGraphClient;
}
