import { BaseLayout } from "@pnp/modern-search-extensibility";
import { IPropertyPaneField, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { Globals, Language } from "../Globals";

export interface INewsArticleLayoutProperties {
    selectedLanguage: string;
    isSearchPage: boolean;
}

export enum NewsArticlePropertyPaneProps {
    SelectedLanguage = 'layoutProperties.selectedLanguage',
    isSearchPage = 'layoutProperties.isSearchPage'
}

export class NewsArticleLayout extends BaseLayout<INewsArticleLayoutProperties> {

    public onInit(): void {
        this.properties.selectedLanguage = this.properties.selectedLanguage !== null ? this.properties.selectedLanguage : Language.English;
        Globals.setLanguage(this.properties.selectedLanguage);

        Globals.setNewsSearchLayout(this.properties.isSearchPage);
    }

    public getPropertyPaneFieldsConfiguration(availableFields: string[]): IPropertyPaneField<any>[] {
        return [
            PropertyPaneTextField(NewsArticlePropertyPaneProps.SelectedLanguage , {
                label: 'Selected language',
                value: Globals.getLanguage(),
                placeholder: `en or fr`
            }),
            PropertyPaneToggle(NewsArticlePropertyPaneProps.isSearchPage, {
                label: 'Search Layout',
                onText: 'We ARE on the search news layout',
                offText: 'We are NOT on the search news layout',
                checked: Globals.getNewsSearchLayout()
            })
        ];
    }

    public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any): void {
        switch (propertyPath) {
            case NewsArticlePropertyPaneProps.SelectedLanguage:
                Globals.setLanguage(newValue);
                break;
            case NewsArticlePropertyPaneProps.isSearchPage:
                Globals.setNewsSearchLayout(newValue);
                break;
        }
    }
}