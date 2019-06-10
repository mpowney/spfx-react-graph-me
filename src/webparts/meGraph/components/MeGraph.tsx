import * as React from 'react';
import { Person } from '@microsoft/microsoft-graph-types';
import { Shimmer } from 'office-ui-fabric-react/lib/Shimmer';

import styles from './styles/MeGraph.module.scss';
import { IMeGraphProps } from './IMeGraphProps';
import PersonTemplate from './templates/Person';
import PersonShimmer from './shimmers/PersonShimmer';

export default class MeGraph extends React.Component<IMeGraphProps, {}> {
    public render(): React.ReactElement<IMeGraphProps> {

        console.log(`MeGraph.tsx: data.value.length = ${this.props.graphData.value.length}`);
        return (
            <div className={ styles.meGraph }>
                <div className={ styles.container }>
                    { this.props.graphData && this.props.selectedEndpoint === '/people' && 
                        this.props.graphData.value.map((person: Person) => {
                            return (
                                <Shimmer 
                                        customElementsGroup={React.createElement(PersonShimmer, {})} 
                                        width={300} 
                                        isDataLoaded={!this.props.isLoading}>
                                    <PersonTemplate {...person} graphClient={this.props.graphClient } />
                                </Shimmer>
                            );
                        })
                    }
                </div>
            </div>
        );
    }

}

