import * as React from 'react';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

import * as strings from 'PropertyFieldGraphPeopleStrings';

import styles from './styles/PropertyFieldGraphPeopleHost.module.scss';
import IGraphPeopleSettings, { ShowUser } from './IGraphPeopleSettings';

export interface IPropertyFieldGraphPeopleHostProps {
	label: string;
	disabled: boolean;
	isHidden: boolean;
	iconName: string;
	onValueChanged: (value: IGraphPeopleSettings) => void;
    value: IGraphPeopleSettings;
}

export interface IPropertyFieldGraphPeopleHostState {
	errorMessage?: string;
    value: IGraphPeopleSettings;
}

export default class PropertyFieldGraphPeopleHost extends React.Component<IPropertyFieldGraphPeopleHostProps, IPropertyFieldGraphPeopleHostState> {

    constructor(props: IPropertyFieldGraphPeopleHostProps, state: IPropertyFieldGraphPeopleHostState) {
		super(props);

		this.state = {
            errorMessage: undefined,
            value: props.value || { ShowUser: ShowUser.CurrentUser, SpecifiedUsername: '' }
		};

    }

    private onChangeShowUser(element?: IDropdownOption): void {
        let updateValue = this.state.value;
        updateValue.ShowUser = ShowUser[element.key];
        this.setState({
            value: updateValue
        }, () => {
            this.props.onValueChanged(updateValue);
        });
      }
    
    
	public render(): JSX.Element {
		return (
			<div className={`${styles.graphPeopleHost} ${this.props.isHidden ? styles.hidden : ""}`}>
                <Label>{this.props.label}</Label>
                <Dropdown label={strings.ShowUserTitle}
                    ariaLabel={strings.ShowUserTitle}
                    defaultSelectedKey={this.props.value && ShowUser[this.props.value.ShowUser]}
                    selectedKey={this.state.value && ShowUser[this.state.value.ShowUser]}
                    options={[
                        { key: ShowUser[ShowUser.CurrentUser], text: strings.ShowUserTitleCurrentUser },
                        { key: ShowUser[ShowUser.SpecifiedUser], text: strings.ShowUserTitleSpecifiedUser }
                    ]}
                    onChanged={this.onChangeShowUser.bind(this)}
                />

            </div>
        );
    }

}
