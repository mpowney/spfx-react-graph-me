import { IPropertyPaneCustomFieldProps } from '@microsoft/sp-webpart-base';

import IGraphPeopleSettings from './IGraphPeopleSettings';

/**
 * Public properties of the PropertyFieldGraphPeople custom field
 */

export interface IPropertyFieldGraphPeopleProps {
    
    /**
    * Property field label displayed on top
    */
    label: string;

    /**
     * Defines an onPropertyChange function to raise when the selected value changes.
     * Normally this function must be defined with the 'this.onPropertyChange'
     * method of the web part object.
     */
    onPropertyChange?(propertyPath: string, oldValue: any, newValue: any): void;

    /**
    * Whether the property pane field is enabled or not.
    */
    disabled?: boolean;

    /**
    * Whether the property pane field is hidden or not.
    */
    isHidden?: boolean;

    /**
     * An UNIQUE key indicates the identity of this control
     */
    //key: string;

    /**
     * Parent Web Part properties
     */
    properties: any;

    /**
     * The name of the UI Fabric Font Icon to use for Inline display (defaults to Color)
     */
    iconName?: string;

    /**
     * The settings held by this property field
     */
    value?: IGraphPeopleSettings;
}

export interface IPropertyFieldGraphPeoplePropsInternal extends IPropertyFieldGraphPeopleProps, IPropertyPaneCustomFieldProps {
}

