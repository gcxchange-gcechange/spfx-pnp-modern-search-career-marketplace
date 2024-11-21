import { BaseLayout } from "@pnp/modern-search-extensibility";
import { IPropertyPaneField, PropertyPaneTextField } from '@microsoft/sp-property-pane';

/**
 * Custom Layout properties
 */
export interface ICustomLayoutProperties {
    selectedLanguage: string;
}

export class CustomLayout extends BaseLayout<ICustomLayoutProperties> {

    public getPropertyPaneFieldsConfiguration(availableFields: string[]): IPropertyPaneField<any>[] {

        this.properties.selectedLanguage = this.properties.selectedLanguage !== null ? this.properties.selectedLanguage : "en";
 
        return [
            PropertyPaneTextField('layoutProperties.selectedLanguage' , {
                label: 'Selected language',
                placeholder: '\'en\' or \'fr\''
            })
        ];
    }
}
