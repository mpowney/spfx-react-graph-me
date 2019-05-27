import * as React from 'react';
import styles from '../styles/Person.module.scss';
import { IPersonaSharedProps, IPersonaProps, Persona } from 'office-ui-fabric-react/lib/Persona';
import { Person } from '@microsoft/microsoft-graph-types';
import { MSGraphClient } from '@microsoft/sp-http';
import { GraphError } from '@microsoft/microsoft-graph-client/lib/src/common';
import { SPComponentLoader } from "@microsoft/sp-loader";
import { ServiceScope } from '@microsoft/sp-core-library';

const COMPONENT_ID: any = {
    SP_WEBPART_SHARED: "914330ee-2df2-4f6e-a858-30c23a812408",
    SP_RTE: "8404d628-4817-4b3a-883e-1c5a4d07892e",
    SP_SUITE_NAV: "f8a8ad94-4cf3-4a19-a76b-1cec9da00219",
    SP_COMPONENT_UTILITIES: "8494e7d7-6b99-47b2-a741-59873e42f16f",
    ODSP_UTILITIES_BUNDLE: "cc2cc925-b5be-41bb-880a-f0f8030c6aff",
    SP_PAGE_CONTEXT: "1c4541f7-5c31-41aa-9fa8-fbc9dc14c0a8"
}

export interface IPersonProps extends Person {
    graphClient: MSGraphClient;
    serviceScope: ServiceScope;
}

export interface IPersonState {
    profilePictureData: any;
    personaCard: any;
}

export default class PersonTemplate extends React.Component<IPersonProps, IPersonState> {

    public constructor(props: IPersonProps) {
        super(props);

        this.state = {
            profilePictureData: null,
            personaCard: null
        };

    }

    public componentDidMount(): void {

        if (this.props.graphClient) {

            const upn = this.props.userPrincipalName;

            const graphPromise = this.props.graphClient
                .api(`/users/${encodeURIComponent(upn)}/photos/48x48/$value`)
                .responseType('blob')
                .get((error: GraphError, response: any, rawResponse?: any) => {
                    // tslint:disable-next-line:no-unused-expression
                    !error.code && response && this.setState({
                        profilePictureData: URL.createObjectURL(response)
                    });

                });

            const componentCardPromise = this.loadSPComponentById(COMPONENT_ID.SP_WEBPART_SHARED).then((sharedLibrary: any) => {
                console.log('sharedLibrary:');
                console.log(sharedLibrary);
                const livePersonaCard: any = sharedLibrary.LivePersonaCard;
                livePersonaCard && this.setState({
                    personaCard: sharedLibrary.LivePersonaCard
                });
            });

            Promise.all([graphPromise, componentCardPromise]).then(promiseResponses => {
            })
        
        }
    }

    private livePersonaCard(container) {
        return React.createElement(this.state.personaCard, {
            className: 'people',
            clientScenario: "PeopleWebPart",
            disableHover: false,
            hostAppPersonaInfo: {
                PersonaType: "User"
            },
            serviceScope: this.props.serviceScope,
            upn: this.cleanUpnFromExternalAddresses(this.props.userPrincipalName),
            name: 'Test Account',
            onCardOpen: () => {

            },
            onCardClose: () => {

            }
        }, container);
    }

    public render(): React.ReactElement<IPersonProps> {

        const personaProps: IPersonaSharedProps = {
            imageInitials: `${this.props.givenName ? this.props.givenName.substring(0, 1) : ''}${this.props.surname ? this.props.surname.substring(0, 1) : ''}`,
            text: this.props.displayName,
            secondaryText: this.props.jobTitle
        };
      
        const livePersonaContainer = <Persona {...personaProps}
                                        onRenderCoin={this._onRenderCoin} />;
        const livePersonaCard = this.state.personaCard && this.livePersonaCard(livePersonaContainer);
        //const livePersonaCard = this.state.webPartTitle && this.webPartTitle(livePersonaContainer);

        return (
            <div className={ styles.person }>
                <div className={ styles.container } style={{color: '#ccc'}}>
                    {livePersonaCard}
                </div>
            </div>
        );
    }

    private _onRenderCoin = (props: IPersonaProps): JSX.Element => {
        const { coinSize, imageAlt, imageUrl } = props;
        return (
            <div className={ styles.coin }>
                { this.state.profilePictureData &&
                    <img src={this.state.profilePictureData} alt={this.props.displayName} width={coinSize} height={coinSize} />
                }
                
            </div>
        );
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
                console.error(`Person.tsx loadSPComponentById(${componentId} SPComponentLoader error: `, error);
                resolve(null);
            });
        });

        return w.cachedspComponent[componentId];
    }

    private cleanUpnFromExternalAddresses(upn: string): string {
        //mark.powney_engagesq.com#ext#@onenil.onmicrosoft.com
        const externalIndex = upn.indexOf('#EXT#');
        return externalIndex > 0 ? upn.substr(0, externalIndex).replace("_", "@") : upn;
    }


}
