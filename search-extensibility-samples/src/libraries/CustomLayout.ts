import { BaseLayout } from "@pnp/modern-search-extensibility";
import { IPropertyPaneField, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { Globals, Language } from "./Globals";

/**
 * Custom Layout properties
 */
export interface ICustomLayoutProperties {
    selectedLanguage: string;
    jobOpportunityPageUrl: string;
}

export enum PropertyPaneProps {
    SelectedLanguage = 'layoutProperties.selectedLanguage',
    JobOpportunityPageUrl = 'layoutProperties.jobOpportunityPageUrl'
}

export class CustomLayout extends BaseLayout<ICustomLayoutProperties> {

    public onInit(): void {
        this.properties.selectedLanguage = this.properties.selectedLanguage !== null ? this.properties.selectedLanguage : Language.English;
        Globals.setLanguage(this.properties.selectedLanguage);
        Globals.jobOpportunityPageUrl = this.properties.jobOpportunityPageUrl;
    }

    public getPropertyPaneFieldsConfiguration(availableFields: string[]): IPropertyPaneField<any>[] {
        return [
            PropertyPaneTextField(PropertyPaneProps.SelectedLanguage , {
                label: 'Selected language',
                value: Globals.getLanguage(),
                placeholder: `${Language.English} or ${Language.French}`
            }),
            PropertyPaneTextField(PropertyPaneProps.JobOpportunityPageUrl, {
                label: 'Job opportunity page URL',
                value: Globals.jobOpportunityPageUrl,
                description: 'Enter the URL for the Job Opportunity page up until where the ID would be.',
                placeholder: 'https://devgcx.sharepoint.com/sites/CM-test/SitePages/Job-Opportunity.aspx?JobOpportunityId='
            })
        ];
    }

    public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any): void {
        switch (propertyPath) {
            case PropertyPaneProps.SelectedLanguage:
                Globals.setLanguage(newValue);
                break;
            case PropertyPaneProps.JobOpportunityPageUrl:
                Globals.jobOpportunityPageUrl = newValue;
                break;
        }
    }
}
