import * as React from 'react';
import * as ReactDom from 'react-dom';

import { IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-webpart-base';

import PropertyFieldColorPickerHost, { IPropertyFieldGraphPeopleHostProps } from './PropertyFieldGraphPeopleHost';
import IGraphPeopleSettings from './IGraphPeopleSettings';
import { IPropertyFieldGraphPeopleProps, IPropertyFieldGraphPeoplePropsInternal} from './IPropertyFieldGraphPeopleProps';


class PropertyFieldGraphPeopleBuilder implements IPropertyPaneField<IPropertyFieldGraphPeopleProps> {

	//Properties defined by IPropertyPaneField
	public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
	public targetProperty: string;
	public properties: IPropertyFieldGraphPeoplePropsInternal;
	private elem: HTMLElement;
	private changeCB?: (targetProperty?: string, newValue?: any) => void;

    public constructor(_targetProperty: string, _properties: IPropertyFieldGraphPeopleProps) {

		this.targetProperty = _targetProperty;
		this.properties = {
            value: _properties.value,
			key: '123', //_properties.key,
			label: _properties.label,
			onPropertyChange: _properties.onPropertyChange,
			disabled: _properties.disabled,
			isHidden: _properties.isHidden,
			properties: _properties.properties,
			iconName: _properties.iconName,
			onRender: this.onRender.bind(this)
		};

    }

	public render(): void {
		if (!this.elem) {
			return;
		}

		this.onRender(this.elem);
	}

	private onRender(elem: HTMLElement, ctx?: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
		if (!this.elem) {
			this.elem = elem;
		}
		this.changeCB = changeCallback;

		const element: React.ReactElement<IPropertyFieldGraphPeopleHostProps> = React.createElement(PropertyFieldColorPickerHost, {
			label: this.properties.label,
			disabled: this.properties.disabled,
			isHidden: this.properties.isHidden,
			iconName: this.properties.iconName || 'Color',
            onValueChanged: this.onValueChanged.bind(this),
            value: this.properties.properties[this.targetProperty]
		});
		ReactDom.render(element, elem);
	}

    private onValueChanged(newValue: IGraphPeopleSettings): void {
        console.log(`onValueChanged() this.targetProperty = ${this.targetProperty}`);
        console.log(`onValueChanged() newValue:`);
        console.log(newValue);
		if (newValue !== null) {
            let oldValue: IGraphPeopleSettings = this.properties.value;
            if (this.properties.onPropertyChange) {
                this.properties.onPropertyChange(this.targetProperty, oldValue, newValue);
            }
			this.properties.properties[this.targetProperty] = newValue;
			if (typeof this.changeCB !== 'undefined' && this.changeCB !== null) {
				this.changeCB(this.targetProperty, newValue);
			}
		}
	}

}

export function PropertyFieldGraphPeople(targetProperty: string, properties: IPropertyFieldGraphPeopleProps): IPropertyPaneField<IPropertyFieldGraphPeopleProps> {
	return new PropertyFieldGraphPeopleBuilder(targetProperty, properties);
}
  