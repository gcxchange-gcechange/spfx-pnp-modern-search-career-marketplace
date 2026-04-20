import { BaseLayout } from "@pnp/modern-search-extensibility";
import { IPropertyPaneField, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { Globals, Language } from "../Globals";

export interface INewsArticleLayoutProperties {
    selectedLanguage: string;
}

export enum NewsArticlePropertyPaneProps {
    SelectedLanguage = 'layoutProperties.selectedLanguage'
}

export class NewsArticleLayout extends BaseLayout<INewsArticleLayoutProperties> {

    public onInit(): void {
        this.properties.selectedLanguage = this.properties.selectedLanguage !== null ? this.properties.selectedLanguage : Language.English;
        Globals.setLanguage(this.properties.selectedLanguage);
    }

    public getPropertyPaneFieldsConfiguration(availableFields: string[]): IPropertyPaneField<any>[] {
        return [
            PropertyPaneTextField(NewsArticlePropertyPaneProps.SelectedLanguage , {
                label: 'Selected language',
                value: Globals.getLanguage(),
                placeholder: `en or fr`
            })
        ];
    }

    public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any): void {
        switch (propertyPath) {
            case NewsArticlePropertyPaneProps.SelectedLanguage:
                Globals.setLanguage(newValue);
                break;
        }
    }
}