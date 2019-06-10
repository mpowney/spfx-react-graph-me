import * as React from 'react';
import styles from '../styles/PersonTemplate.module.scss';
import { IPersonaSharedProps, IPersonaProps, Persona } from 'office-ui-fabric-react/lib/Persona';
import { Person } from '@microsoft/microsoft-graph-types';
import { MSGraphClient } from '@microsoft/sp-http';
import { GraphError } from '@microsoft/microsoft-graph-client/lib/src/common';

export interface IPersonTemplateProps extends Person {
    graphClient: MSGraphClient;
}

export interface IPersonTemplateState {
    profilePictureData: any;
}

export default class PersonTemplate extends React.Component<IPersonTemplateProps, IPersonTemplateState> {

    public constructor(props: IPersonTemplateProps) {
        super(props);

        this.state = {
            profilePictureData: null
        };

    }

    public componentDidMount(): void {

        if (this.props.graphClient) {

            const upn = this.props.userPrincipalName;

            this.props.graphClient
                .api(`/users/${encodeURIComponent(upn)}/photos/48x48/$value`)
                .responseType('blob')
                .get((error: GraphError, response: any, rawResponse?: any) => {
                    // tslint:disable-next-line:no-unused-expression
                    !error.code && response && this.setState({
                        profilePictureData: URL.createObjectURL(response)
                    });

                });

        }
    }

    public render(): React.ReactElement<IPersonTemplateProps> {

        const personaProps: IPersonaSharedProps = {
            imageInitials: `${this.props.givenName ? this.props.givenName.substring(0, 1) : ''}${this.props.surname ? this.props.surname.substring(0, 1) : ''}`,
            text: this.props.displayName,
            secondaryText: this.props.jobTitle
        };
      

        return (
            <div className={ styles.person }>
                <div className={ styles.container }>
                    <Persona {...personaProps}
                        onRenderCoin={this._onRenderCoin} />
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
}
