import * as React from 'react';
import styles from './styles/MeGraph.module.scss';
import { IMeGraphProps } from './IMeGraphProps';
import { Person } from '@microsoft/microsoft-graph-types';

import PersonTemplate from './templates/Person';

export default class MeGraph extends React.Component<IMeGraphProps, {}> {
  public render(): React.ReactElement<IMeGraphProps> {
    return (
      <div className={ styles.meGraph }>
        <div className={ styles.container }>
          { this.props.graphData && this.props.selectedEndpoint === '/people' && 
              this.props.graphData.value.map((person: Person) => {
                return <PersonTemplate {...person} graphClient={this.props.graphClient } serviceScope={this.props.serviceScope} />;
            })
          }
        </div>
      </div>
    );
  }
}
