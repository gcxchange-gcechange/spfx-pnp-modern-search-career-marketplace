import { BaseLayout } from "@pnp/modern-search-extensibility";
import { IPropertyPaneField, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { Globals, Language } from "./Globals";

export interface ICustomLayoutProperties {
    selectedLanguage: string;
    jobOpportunityPageUrl: string;
    applicationDeadlineDate: string;
    cityEn: string;
    cityFr: string;
    classificationLevel: string;
    contactEmail: string;
    contactName: string;
    contactObjectId: string;
    durationEn: string;
    durationFr: string;
    jobDescriptionEn: string;
    jobDescriptionFr: string;
    jobTitleEn: string;
    jobTitleFr: string;
    jobType: string;
    durationQuantity: string;
    jobTypeTermSetGuid: string;
}

export enum PropertyPaneProps {
    SelectedLanguage = 'layoutProperties.selectedLanguage',
    JobOpportunityPageUrl = 'layoutProperties.jobOpportunityPageUrl',
    ApplicationDeadlineDate = 'layoutProperties.applicationDeadlineDate',
    CityEn = 'layoutProperties.cityEn',
    CityFr = 'layoutProperties.cityFr',
    ClassificationLevel = 'layoutProperties.classificationLevel',
    ContactEmail = 'layoutProperties.contactEmail',
    ContactName = 'layoutProperties.contactName',
    ContactObjectId = 'layoutProperties.contactObjectId',
    DurationEn = 'layoutProperties.durationEn',
    DurationFr = 'layoutProperties.durationFr',
    JobDescriptionEn = 'layoutProperties.jobDescriptionEn',
    JobDescriptionFr = 'layoutProperties.jobDescriptionFr',
    JobTitleEn = 'layoutProperties.jobTitleEn',
    JobTitleFr = 'layoutProperties.jobTitleFr',
    JobType = 'layoutProperties.jobType',
    DurationQuantity = 'layoutProperties.durationQuantity',
    JobTypeTermSetGuid = 'layoutProperties.jobTypeTermSetGuid'
}

export class CustomLayout extends BaseLayout<ICustomLayoutProperties> {

    public onInit(): void {
        this.properties.selectedLanguage = this.properties.selectedLanguage !== null ? this.properties.selectedLanguage : Language.English;
        Globals.setLanguage(this.properties.selectedLanguage);
        Globals.jobOpportunityPageUrl = this.properties.jobOpportunityPageUrl;
        this.getJobTypes();
    }

    private validateRequiredField(value: string): string {
        return value && value.trim().length > 0 ? '' : 'This field is required.';
    }

    public getPropertyPaneFieldsConfiguration(availableFields: string[]): IPropertyPaneField<any>[] {
        return [
            PropertyPaneTextField(PropertyPaneProps.SelectedLanguage , {
                label: 'Selected language',
                value: Globals.getLanguage(),
                placeholder: `en or fr`
            }),
            PropertyPaneTextField(PropertyPaneProps.JobOpportunityPageUrl, {
                label: 'Job opportunity page URL',
                value: Globals.jobOpportunityPageUrl,
                description: 'Enter the URL for the Job Opportunity page up until where the ID would be.',
                placeholder: 'https://devgcx.sharepoint.com/sites/CM-test/SitePages/Job-Opportunity.aspx?JobOpportunityId=',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.ApplicationDeadlineDate, {
                label: 'ApplicationDeadlineDate Managed Property',
                value: this.properties.applicationDeadlineDate,
                placeholder: 'CM-ApplicationDeadlineDate',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.CityEn, {
                label: 'CityEn Managed Property',
                value: this.properties.cityEn,
                placeholder: 'CM-City',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.CityFr, {
                label: 'CityFr Managed Property',
                value: this.properties.cityFr,
                placeholder: 'CM-CityFr',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.ClassificationLevel, {
                label: 'ClassificationLevel Managed Property',
                value: this.properties.classificationLevel,
                placeholder: 'CM-ClassificationLevel',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.ContactEmail, {
                label: 'ContactEmail Managed Property',
                value: this.properties.contactEmail,
                placeholder: 'CM-ContactEmail',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.ContactName, {
                label: 'ContactName Managed Property',
                value: this.properties.contactName,
                placeholder: 'CM-ContactName',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.ContactObjectId, {
                label: 'ContactObjectId Managed Property',
                value: this.properties.contactObjectId,
                placeholder: 'CM-ContactObjectId',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.DurationEn, {
                label: 'DurationEn Managed Property',
                value: this.properties.durationEn,
                placeholder: 'CM-Duration',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.DurationFr, {
                label: 'DurationFr Managed Property',
                value: this.properties.durationFr,
                placeholder: 'CM-DurationFr',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.JobDescriptionEn, {
                label: 'JobDescriptionEn Managed Property',
                value: this.properties.jobDescriptionEn,
                placeholder: 'CM-JobDescriptionEn',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.JobDescriptionFr, {
                label: 'JobDescriptionFr Managed Property',
                value: this.properties.jobDescriptionFr,
                placeholder: 'CM-JobDescriptionFr',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.JobTitleEn, {
                label: 'JobTitleEn Managed Property',
                value: this.properties.jobTitleEn,
                placeholder: 'CM-JobTitleEn',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.JobTitleFr, {
                label: 'JobTitleFr Managed Property',
                value: this.properties.jobTitleFr,
                placeholder: 'CM-JobTitleFr',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.JobType, {
                label: 'JobType Managed Property',
                value: this.properties.jobType,
                placeholder: 'CM-JobType',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.DurationQuantity, {
                label: 'DurationQuantity Managed Property',
                value: this.properties.durationQuantity,
                placeholder: 'CM-DurationQuantity',
                onGetErrorMessage: this.validateRequiredField.bind(this)
            }),
            PropertyPaneTextField(PropertyPaneProps.JobTypeTermSetGuid, {
                label: 'JobType term set GUID',
                value: this.properties.jobTypeTermSetGuid,
                placeholder: '45f37f08-3ff4-4d84-bf21-4a77ddffcf3e',
                onGetErrorMessage: this.validateRequiredField.bind(this)
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

    private async getJobTypes() {
        try {
            const response = await fetch(`/_api/v2.1/termstore/sets/${this.properties.jobTypeTermSetGuid}/terms/`, {
                method: 'GET',
                headers: { 'Accept': 'application/json;odata=verbose' }
            });
            
            if (!response.ok) throw new Error(`Failed to fetch term set: ${this.properties.jobTypeTermSetGuid}`);
            
            const jobTypes = await response.json();

            Globals.setJobTypes(jobTypes.value)
        } catch (error) {
            console.error("Error fetching term set:", error);
        }
    }
}
