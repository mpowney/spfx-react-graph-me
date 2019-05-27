import { MSGraphClient } from '@microsoft/sp-http';
import { ServiceScope, DisplayMode } from '@microsoft/sp-core-library';

export interface IMeGraphProps {
    selectedEndpoint: string;
    graphData: any;
    isLoading: boolean;
    graphClient: MSGraphClient;
    serviceScope: ServiceScope;
}
