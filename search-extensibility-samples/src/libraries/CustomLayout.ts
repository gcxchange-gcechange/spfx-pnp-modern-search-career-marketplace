import { BaseLayout } from "@pnp/modern-search-extensibility";
import { IPropertyPaneField, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { Globals, Language } from "./Globals";

/**
 * Custom Layout properties
 */
export interface ICustomLayoutProperties {
    selectedLanguage: string;
}

export enum PropertyPaneProps {
    SelectedLanguage = 'layoutProperties.selectedLanguage'
}

export class CustomLayout extends BaseLayout<ICustomLayoutProperties> {

    public onInit(): void {
        this.properties.selectedLanguage = this.properties.selectedLanguage !== null ? this.properties.selectedLanguage : Language.English;
        Globals.setLanguage(this.properties.selectedLanguage);
    }

    public getPropertyPaneFieldsConfiguration(availableFields: string[]): IPropertyPaneField<any>[] {
        return [
            PropertyPaneTextField(PropertyPaneProps.SelectedLanguage , {
                label: 'Selected language',
                value: Globals.getLanguage(),
                placeholder: `${Language.English} or ${Language.French}`
            })
        ];
    }

    public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any): void {
        switch (propertyPath) {
            case PropertyPaneProps.SelectedLanguage:
                Globals.setLanguage(newValue);
                break;
        }
    }
}
