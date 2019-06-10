import * as React from 'react';
import { ShimmerElementsGroup, ShimmerElementType } from 'office-ui-fabric-react/lib/Shimmer';

import styles from '../styles/PersonShimmer.module.scss';

export default class PersonShimmerTemplate extends React.Component<{}, {}> {

    public constructor(props: {}) {
        super(props);

        this.state = {
        };

    }

    public componentDidMount(): void {
    }

    public render(): React.ReactElement<{}> {
        return (
            <div className={ styles.personShimmer }>

                <ShimmerElementsGroup
                    shimmerElements={[{ type: ShimmerElementType.circle, height: 40 }, { type: ShimmerElementType.gap, width: 16, height: 40 }]}
                />
                <ShimmerElementsGroup
                    flexWrap={true}
                    width="100%"
                    shimmerElements={[
                        { type: ShimmerElementType.line, width: '100%', height: 10, verticalAlign: 'bottom' },
                        { type: ShimmerElementType.line, width: '90%', height: 8 },
                        { type: ShimmerElementType.gap, width: '10%', height: 20 }
                    ]}
                />
            </div>
        );

    }

}