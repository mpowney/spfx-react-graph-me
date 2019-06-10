import * as React from 'react';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';

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
            value: props.value || { 
                ShowUser: ShowUser.CurrentUser, 
                SpecifiedUsername: '',
                FilterOnlyUsers: false
            }
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

    private specifiedUserTimeout: any;
    private onChangeSpecifiedUser(value: string): void {
        let updateValue = this.state.value;
        updateValue.SpecifiedUsername = value;
        this.setState({
            value: updateValue
        }, () => {
            if (this.specifiedUserTimeout) {
                window.clearTimeout(this.specifiedUserTimeout);
            }
            this.specifiedUserTimeout = window.setTimeout(() => {
                this.props.onValueChanged(updateValue);
            }, 1000);
        });
    }

    private onChangeFilterOnlyUsers(value: boolean): void {
        let updateValue = this.state.value;
        updateValue.FilterOnlyUsers = value;
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
                { this.state.value && ShowUser[this.state.value.ShowUser] === ShowUser[ShowUser.SpecifiedUser] &&
                    <TextField label={strings.SpecifiedUserLabel} 
                        ariaLabel={strings.SpecifiedUserLabel}
                        defaultValue={this.props.value.SpecifiedUsername}
                        value={this.state.value.SpecifiedUsername}
                        onChanged={this.onChangeSpecifiedUser.bind(this)}
                    />
                }
                <Toggle label={strings.FilterOnlyUsersLabel}
                    ariaLabel={strings.FilterOnlyUsersLabel}
                    defaultChecked={this.props.value.FilterOnlyUsers}
                    checked={this.state.value.FilterOnlyUsers}
                    onChanged={this.onChangeFilterOnlyUsers.bind(this)}
                />
            </div>
        );
    }

}
