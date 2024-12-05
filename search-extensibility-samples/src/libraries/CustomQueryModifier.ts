import { BaseQueryModifier } from "@pnp/modern-search-extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as myLibraryStrings from 'MyCompanyLibraryLibraryStrings';
import { Globals, Language } from "./Globals";

export interface IAdvancedSearchQueryModifierProperties {
  searchBoxSelector: string;
  searchButtonId: string;
  clearButtonId: string;
  jobTitleId: string;
  departmentId: string;
  classificationCodeId: string;
  classificationLevelId: string;
  languageRequirementId: string;
  regionId: string;
  durationId: string;
}

//CustomQueryModifier
export class AdvancedSearchQueryModifier extends BaseQueryModifier<IAdvancedSearchQueryModifierProperties> {

  private static readonly DEFAULT_VALUE = '*';
  private static readonly SESSION_STORAGE_KEYS = [
    'gcx-cm-jobTitle',
    'gcx-cm-departmentId',
    'gcx-cm-classificationCodeId',
    'gcx-cm-classificationLevelId',
    'gcx-cm-languageRequirement',
    'gcx-cm-regionId',
    'gcx-cm-durationId'
  ];

  public async onInit(): Promise<void> {
    debugger;
    AdvancedSearchQueryModifier.SESSION_STORAGE_KEYS.forEach(key => {
      if (!sessionStorage.getItem(key)) {
        sessionStorage.setItem(key, AdvancedSearchQueryModifier.DEFAULT_VALUE);
      }
    });
    this.setupListeners();
  }

  private setupListeners(): void {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const context = this;

    if (this._properties.searchButtonId) {
      let el = document.getElementById(this._properties.searchButtonId);
      if (el) {
        el.addEventListener('click', (event) => {
          setTimeout(() => {
            context.triggerSearch();
          }, 0);
        });
      } else { console.error(`Advanced Search: Couldn't find advanced search button element with the ID \'${this._properties.searchButtonId}\'`); }
    } else { console.error(`Advanced Search: No ID provided for SearchButton`); }

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

            if (searchBox.defaultValue === "") {
                searchBox.value = " ";
                searchBox.dispatchEvent(new Event('input', { bubbles: true }));
            }

            el.dispatchEvent(new KeyboardEvent('keydown', {
                key: 'Enter',
                code: 'Enter',
                keyCode: 13,
                which: 13,
                bubbles: true,
                cancelable: true,
            }));
        }
    }
  }

  public async modifyQuery(queryText: string): Promise<string> {
    queryText = queryText || AdvancedSearchQueryModifier.DEFAULT_VALUE;

    const jobTitle = sessionStorage.getItem('gcx-cm-jobTitle') || AdvancedSearchQueryModifier.DEFAULT_VALUE;
    const department = sessionStorage.getItem('gcx-cm-departmentId') || AdvancedSearchQueryModifier.DEFAULT_VALUE;
    const classificationCode = sessionStorage.getItem('gcx-cm-classificationCodeId') || AdvancedSearchQueryModifier.DEFAULT_VALUE;
    const classificationLevel = sessionStorage.getItem('gcx-cm-classificationLevelId') || AdvancedSearchQueryModifier.DEFAULT_VALUE;
    const languageRequirement = sessionStorage.getItem('gcx-cm-languageRequirementId') || AdvancedSearchQueryModifier.DEFAULT_VALUE;
    // TODO: Figure out how we are doing location/region
    const region = sessionStorage.getItem('gcx-cm-regionId') || AdvancedSearchQueryModifier.DEFAULT_VALUE;
    const duration = sessionStorage.getItem('gcx-cm-durationId') || AdvancedSearchQueryModifier.DEFAULT_VALUE;

    if (Globals.getLanguage() === Language.French) {
      return `${queryText} path: https://devgcx.sharepoint.com/sites/CM-test/Lists/JobOpportunity/ contentclass: STS_ListItem_GenericList "CM-JobTitleFr":*${jobTitle}* AND "CM-LanguageRequirementId":${languageRequirement} AND "CM-DepartmentId":${department} AND "CM-ClassificationCodeId":${classificationCode} AND "CM-ClassificationLevelId":${classificationLevel} AND "CM-DurationId":${duration}`;
    } else {
      return `${queryText} path: https://devgcx.sharepoint.com/sites/CM-test/Lists/JobOpportunity/ contentclass: STS_ListItem_GenericList "CM-JobTitleEn":*${jobTitle}* AND "CM-LanguageRequirementId":${languageRequirement} AND "CM-DepartmentId":${department} AND "CM-ClassificationCodeId":${classificationCode} AND "CM-ClassificationLevelId":${classificationLevel} AND "CM-DurationId":${duration}`;
    }
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
          })
        ],
      },
    ];
  }
}