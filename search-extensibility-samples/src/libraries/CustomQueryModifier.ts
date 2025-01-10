import { BaseQueryModifier } from "@pnp/modern-search-extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as myLibraryStrings from 'MyCompanyLibraryLibraryStrings';
import { Globals, Language } from "./Globals";

export interface IAdvancedSearchQueryModifierProperties {
  listPath: string;
  searchBoxSelector: string;
  searchButtonId: string;
  clearButtonId: string;
  jobTitleEnMP: string;
  jobTitleFrMP: string;
  departmentMP: string;
  classificationCodeMP: string;
  classificationLevelMP: string;
  languageRequirementMP: string;
  regionMP: string;
  durationMP: string;
}

enum AdvancedSearchSessionKeys {
  JobTitle = 'gcx-cm-jobTitle',
  ClassificationCode = 'gcx-cm-classificationCode',
  ClassificationLevel = 'gcx-cm-classificationLevel',
  Department = 'gcx-cm-department',
  Duration = 'gcx-cm-duration',
  LanguageRequirement = 'gcx-cm-languageRequirement',
  Location = 'gcx-cm-location',
}

//CustomQueryModifier
export class AdvancedSearchQueryModifier extends BaseQueryModifier<IAdvancedSearchQueryModifierProperties> {
  private static readonly DEFAULT_VALUE = '*';

  public async onInit(): Promise<void> {

    // Initialize the session storage items
    (Object.keys(AdvancedSearchSessionKeys) as (keyof typeof AdvancedSearchSessionKeys)[]).forEach(key => {
      const value = AdvancedSearchSessionKeys[key];
      if (!sessionStorage.getItem(value)) {
        sessionStorage.setItem(value, '');
      }
    });

    this.setupListeners();
  }

  private setupListeners(): void {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const context = this;
    
    // Search button
    if (this._properties.searchButtonId) {
      let el = document.getElementById(this._properties.searchButtonId);
      if (el) {
        el.addEventListener('click', (event) => {
          event.preventDefault();
          setTimeout(() => {
            context.triggerSearch();
          }, 0);
        });
      } else { console.error(`Advanced Search: Couldn't find advanced search button element with the ID \'${this._properties.searchButtonId}\'`); }
    } else { console.error(`Advanced Search: No ID provided for SearchButton`); }

    // Clear button
    if (this._properties.clearButtonId) {
      let el = document.getElementById(this._properties.clearButtonId);
      if (el) {
        el.addEventListener('click', (event) => {
          event.preventDefault();
          setTimeout(() => {
            context.triggerSearch();
          }, 0);
        });
      } else { console.error(`Advanced Search: Couldn't find advanced clear button element with the ID \'${this._properties.clearButtonId}\'`); }
    } else { console.error(`Advanced Search: No ID provided for ClearButton`); }
  }

  private triggerSearch(): void {
    if (this._properties.searchBoxSelector) {
        let el = document.querySelector(this._properties.searchBoxSelector);
        if (el) {
            let searchBox = el as HTMLInputElement;

            // If pnp search box has no value insert a space so the input becomes active
            if (searchBox.defaultValue === "") {
                searchBox.value = " ";
                searchBox.dispatchEvent(new Event('input', { bubbles: true }));
            }

            // Send a enter keydown event to the pnp search box input to perform the search
            el.dispatchEvent(new KeyboardEvent('keydown', {
                key: 'Enter',
                code: 'Enter',
                keyCode: 13,
                which: 13,
                bubbles: true,
                cancelable: true,
            }));
        } else { console.error(`Advanced Search: Couldn't find PnP Search Box Input via selector \'${this._properties.clearButtonId}\'`); }
    }
  }

  public async modifyQuery(queryText: string): Promise<string> {
    queryText = queryText || AdvancedSearchQueryModifier.DEFAULT_VALUE;

    if (queryText.trim() == '')
      queryText = '*';

    let finalQuery = `${queryText} path: ${this._properties.listPath} contentclass: STS_ListItem_GenericList `;
    let propSet = false;

    const jobTitle = sessionStorage.getItem(AdvancedSearchSessionKeys.JobTitle);
    if (jobTitle && jobTitle.trim() != '') {
      if (Globals.getLanguage() === Language.French) {
        finalQuery += `"${this._properties.jobTitleFrMP}":*${jobTitle}* `;
      } else {
        finalQuery += `"${this._properties.jobTitleEnMP}":*${jobTitle}* `;
      }
      propSet = true;
    }

    const department = sessionStorage.getItem(AdvancedSearchSessionKeys.Department);
    if (department && department.trim() != '') {
      finalQuery += `${propSet ? 'AND ' : ''}"${this._properties.departmentMP}":${department} `;
      propSet = true;
    }

    const classificationCode = sessionStorage.getItem(AdvancedSearchSessionKeys.ClassificationCode);
    if (classificationCode && classificationCode.trim() != '') {
      finalQuery += `${propSet ? 'AND ' : ''}"${this._properties.classificationCodeMP}":${classificationCode} `;
      propSet = true;
    }

    const classificationLevel = sessionStorage.getItem(AdvancedSearchSessionKeys.ClassificationLevel);
    if (classificationLevel && classificationLevel.trim() != '') {
      finalQuery += `${propSet ? 'AND ' : ''}"${this._properties.classificationLevelMP}":${classificationLevel} `;
      propSet = true;
    }

    const languageRequirement = sessionStorage.getItem(AdvancedSearchSessionKeys.LanguageRequirement);
    if (languageRequirement && languageRequirement.trim() != '') {
      finalQuery += `${propSet ? 'AND ' : ''}"${this._properties.languageRequirementMP}":${languageRequirement} `;
      propSet = true;
    }

    const duration = sessionStorage.getItem(AdvancedSearchSessionKeys.Duration);
    if (duration && duration.trim() != '') {
      finalQuery += `${propSet ? 'AND ' : ''}"${this._properties.durationMP}":${duration} `;
      propSet = true;
    }

    // TODO: Implement location
    const location = sessionStorage.getItem(AdvancedSearchSessionKeys.Location) || AdvancedSearchQueryModifier.DEFAULT_VALUE;

    return finalQuery;
  }

  // TODO: Update listeners
  public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any): void {
    switch(propertyPath) {
      case 'queryModifierProperties.searchBoxSelector':
        break;
      case 'queryModifierProperties.searchButton':
        break;
      case 'queryModifierProperties.clearButton':
        break;
    }
  }

  public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {

    return [
      {
        groupName: myLibraryStrings.CustomQueryModifier.GroupName,
        groupFields: [
          PropertyPaneTextField('queryModifierProperties.listPath', {
            label: 'JobOpportunity List Path',
            description: 'The path to the JobOpportunity list on the site this webpart is deployed.',
            placeholder: 'https://devgcx.sharepoint.com/sites/CM-test/Lists/JobOpportunity/',
          }),
          PropertyPaneTextField('queryModifierProperties.searchBoxSelector', {
            label: 'PnP Search Box Selector',
            description: 'CSS selector for the pnp search input.',
            placeholder: '[data-sp-feature-tag="pnpSearchBoxWebPart web part (PnP - Search Box)"] input',
          }),
          PropertyPaneTextField('queryModifierProperties.searchButtonId', {
            label: 'Advanced Search - Search Button ID',
            description: 'The ID of the advanced search\'s search button.',
            placeholder: 'advancedSearch-Search',
          }),
          PropertyPaneTextField('queryModifierProperties.clearButtonId', {
            label: 'Advanced Search - Clear Button ID',
            description: 'The ID of the advanced search\'s clear button.', 
            placeholder: 'advancedSearch-Clear',
          }),
          PropertyPaneTextField('queryModifierProperties.jobTitleEnMP', {
            label: 'English JobTitle Managed Property',
            description: 'The managed property name for the English JobTitle', 
            placeholder: 'CM-JobTitleEn',
          }),
          PropertyPaneTextField('queryModifierProperties.jobTitleFrMP', {
            label: 'French JobTitle Managed Property',
            description: 'The managed property name for the French JobTitle', 
            placeholder: 'CM-JobTitleFr',
          }),
          PropertyPaneTextField('queryModifierProperties.departmentMP', {
            label: 'Department Managed Property',
            description: 'The managed property name for Department', 
            placeholder: 'CM-DepartmentId',
          }),
          PropertyPaneTextField('queryModifierProperties.classificationCodeMP', {
            label: 'ClassificationCode Managed Property',
            description: 'The managed property name for ClassificationCode', 
            placeholder: 'CM-ClassificationCodeId',
          }),
          PropertyPaneTextField('queryModifierProperties.classificationLevelMP', {
            label: 'ClassificationLevel Managed Property',
            description: 'The managed property name for ClassificationLevel', 
            placeholder: 'CM-ClassificationLevelId',
          }),
          PropertyPaneTextField('queryModifierProperties.languageRequirementMP', {
            label: 'LanguageRequirement Managed Property',
            description: 'The managed property name for LanguageRequirement', 
            placeholder: 'CM-LanguageRequirementId',
          }),
          PropertyPaneTextField('queryModifierProperties.regionMP', { // TODO: Update to location
            label: 'Region Managed Property',
            description: 'The managed property name for Region', 
            placeholder: 'CM-RegionId',
          }),
          PropertyPaneTextField('queryModifierProperties.durationMP', {
            label: 'Duration Managed Property',
            description: 'The managed property name for Duration', 
            placeholder: 'CM-DurationId',
          })
        ],
      },
    ];
  }
}