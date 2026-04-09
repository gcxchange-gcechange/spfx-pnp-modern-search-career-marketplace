import { BaseQueryModifier } from "@pnp/modern-search-extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as myLibraryStrings from 'MyCompanyLibraryLibraryStrings';
import { Globals, Language } from "./Globals";

export interface IAdvancedSearchQueryModifierProperties {
  listPath: string;
  searchBoxSelector: string;
  searchButtonId: string;
  clearButtonId: string;
  filterButtonId:string;
  clearFilterButtonId: string;
  jobTitleEnMP: string;
  jobTitleFrMP: string;
  departmentMP: string;
  classificationCodeMP: string;
  classificationLevelMP: string;
  languageRequirementMP: string;
  cityMP: string;
  deadlineFilterMP: string;
  jobTypeMP: string;
  workArrangementMP: string;
}

export enum FilterSessionKeys {
  Initialized = 'gcx-cm-filter-init',
  JobType = 'gcx-cm-filter-jobType',
  ClassificationCode = 'gcx-cm-filter-classificationCode',
  ClassificationLevel = 'gcx-cm-filter-clasificationLevel',
  Department = 'gcx-cm-filter-department',
  WorkArrangement = 'gcx-cm-filter-workArrangement',
  City = 'gcx-cm-filter-city',
  LanguageRequirement = 'gcx-cm-filter-languageRequirement'
}

export enum QueryModifierKeys {
  AdvancedSearch = 'gcx-cm-querymod-as',
  MyOpportunities = 'gcx-cm-querymod-myop'
}

//CustomQueryModifier
export class AdvancedSearchQueryModifier extends BaseQueryModifier<IAdvancedSearchQueryModifierProperties> {
  private static readonly DEFAULT_VALUE = '*';
  private lang = Globals.getLanguage();
  private todayIso: string;

  public async onInit(): Promise<void> {

    // Initialize the filter session storage items
    (Object.keys(FilterSessionKeys) as (keyof typeof FilterSessionKeys)[]).forEach(key => {
      const value = FilterSessionKeys[key];
      if (value !== FilterSessionKeys.Initialized)
        sessionStorage.setItem(value, '');
    });

    this.setupListeners();

    const today = new Date();
    today.setUTCHours(0, 0, 0, 0); 
    this.todayIso = today.toISOString();
  }

  private setupListeners(): void {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const context = this;
    let attempts = 0;
    const maxAttempts = 30;
    const attemptInterval = 1000;

    // eslint-disable-next-line @typescript-eslint/ban-types
    const tryGetSessionStorageItem = (key: string, callback: Function, interval: number = 1000, maxAttempts: number = Number.MAX_VALUE): void => {
      let attempts: number = 0;

      const getKey = setInterval(() => {
        const item = sessionStorage.getItem(key);

        if (item !== null) {
          sessionStorage.removeItem(key);
          clearInterval(getKey);
          callback();
        } else {
          attempts++;
          if (attempts >= maxAttempts) {
            clearInterval(getKey);
            console.error(`Query Modifier: Couldn't find sessionStorage item with key '${key}' after ${maxAttempts} attempts over ${maxAttempts * interval / 1000} seconds.`);
          }
        }
      }, interval);
    }

    const tryGetElement = (id: string, callback: any) => {
      if (!id) {
        console.error(`Query Modifier: No ID provided`);
        return;
      }
  
      const interval = setInterval(() => {
        let el = document.getElementById(id);
        if (el) {
          clearInterval(interval);
          callback(el);
        } else {
          attempts++;
          if (attempts >= maxAttempts) {
            clearInterval(interval);
            console.error(`Query Modifier: Couldn't find element with the ID '${id}' after ${maxAttempts} attempts over ${maxAttempts * attemptInterval / 1000} seconds.`);
          }
        }
      }, attemptInterval);
    };

    tryGetSessionStorageItem(FilterSessionKeys.Initialized, () => {
      // Filters - Apply Btn
      tryGetElement(this._properties.filterButtonId, (el: HTMLElement) => {
        el.addEventListener('click', (event) => {
          event.preventDefault();
          setTimeout(() => {
            context.triggerSearch();
          }, 0);
        });
      });
      // Filters - Clear Btn
      tryGetElement(this._properties.clearFilterButtonId, (el: HTMLElement) => {
        el.addEventListener('click', (event) => {
          event.preventDefault();
          setTimeout(() => {
            context.triggerSearch();
          }, 0);
        });
      });
    });
  }

  private triggerSearch(cleanSearch: boolean = false): void {
    if (this._properties.searchBoxSelector) {
        let el = document.querySelector(this._properties.searchBoxSelector);
        if (el) {
            let searchBox = el as HTMLInputElement;

            // pnp search box needs a value so the input becomes active
            if (cleanSearch || searchBox.defaultValue === "") {
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
        } else { console.error(`QUery Modifier: Couldn't find PnP Search Box Input via selector \'${this._properties.clearButtonId}\'`); }
    }
  }

  public async modifyQuery(queryText: string): Promise<string> {
    queryText = queryText || AdvancedSearchQueryModifier.DEFAULT_VALUE;

    if (queryText.trim() == '')
      queryText = '*';

    let finalQuery = this.applyFilters(`${queryText !== '*' ? '*' + queryText + '*' : queryText} path: ${this._properties.listPath} contentclass: STS_ListItem_GenericList`);

    console.log(finalQuery);

    // Set this item so the other custom queries know we've already performed an advanced search/filter on the original query
    sessionStorage.setItem(QueryModifierKeys.AdvancedSearch, 'true');

    return finalQuery;
  }

  private applyFilters(query: string): string {
    const jobTypes = sessionStorage.getItem(FilterSessionKeys.JobType);
    const classCodes = sessionStorage.getItem(FilterSessionKeys.ClassificationCode);
    const classLevels = sessionStorage.getItem(FilterSessionKeys.ClassificationLevel);
    const departments = sessionStorage.getItem(FilterSessionKeys.Department);
    const workArrangements = sessionStorage.getItem(FilterSessionKeys.WorkArrangement);
    const cities = sessionStorage.getItem(FilterSessionKeys.City);
    const languageRequirements = sessionStorage.getItem(FilterSessionKeys.LanguageRequirement);

    if ((jobTypes === undefined || jobTypes.trim() == '') && (classCodes === undefined || classCodes.trim() == '')
     && (classLevels === undefined || classLevels.trim() == '') && (departments === undefined || departments.trim() == '')
     && (workArrangements === undefined || workArrangements.trim() == '') && (cities === undefined || cities.trim() == '')
     && (languageRequirements === undefined || languageRequirements.trim() == '')
    ) {
      return query;
    }

    let finalQuery = `${query} AND (`;

    finalQuery += `"${this._properties.deadlineFilterMP}">=${this.todayIso} `;

    finalQuery = this.AddFilterQuery(finalQuery, this._properties.jobTypeMP, jobTypes);
    finalQuery = this.AddFilterQuery(finalQuery, this._properties.classificationCodeMP, classCodes);
    finalQuery = this.AddFilterQuery(finalQuery, this._properties.classificationLevelMP, classLevels);
    finalQuery = this.AddFilterQuery(finalQuery, this._properties.departmentMP, departments);
    finalQuery = this.AddFilterQuery(finalQuery, this._properties.workArrangementMP, workArrangements);
    finalQuery = this.AddFilterQuery(finalQuery, this._properties.cityMP, cities);
    finalQuery = this.AddFilterQuery(finalQuery, this._properties.languageRequirementMP, languageRequirements);

    return `${finalQuery})`;
  }

  private AddFilterQuery(query: string, managedProperty: string, selections: string, ): string {
    let retVal = query;
    if (selections && selections.trim() != '') {
      const selectionArr = selections.split(',');
      for (let i = 0; i < selectionArr.length; i++) {

        if (i == 0)
          retVal += `AND (`;

        retVal += `"${managedProperty}":${selectionArr[i]}`;

        if (i != selectionArr.length - 1)
          retVal += ' OR ';
      }

      retVal += ')';
    }
    return retVal;
  }

  // private getAllLanguageComprehensions(languageRequirement: string): string[] {
  //   const results: string[] = [];

  //   function helper(current: string, index: number) {
  //       if (index === languageRequirement.length) {
  //           results.push(current);
  //           return;
  //       }

  //       const char = languageRequirement[index];
  //       if (char === '-') {
  //           helper(current + '-', index + 1);
  //       } else if (char === 'A') {
  //           helper(current + 'A', index + 1);
  //           helper(current + 'B', index + 1);
  //           helper(current + 'C', index + 1);
  //       } else if (char === 'B') {
  //           helper(current + 'B', index + 1);
  //           helper(current + 'C', index + 1);
  //       } else if (char === 'C') {
  //           helper(current + 'C', index + 1);
  //       }
  //   }

  //   helper('', 0);
  //   return results;
  // }

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
          PropertyPaneTextField('queryModifierProperties.filterButtonId', {
            label: 'Filter - Apply Button ID',
            description: 'The ID of the button that applies the filters.',
            placeholder: 'gcx-cm-filter-apply',
          }),
          PropertyPaneTextField('queryModifierProperties.clearFilterButtonId', {
            label: 'Filter - Clear Button ID',
            description: 'The ID of the button that clears the filters.',
            placeholder: 'gcx-cm-filter-clear',
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
          // PropertyPaneTextField('queryModifierProperties.languageComprehensionMP', {
          //   label: 'Language Comprehension Managed Property',
          //   description: 'The managed property name for LanguageComprehension', 
          //   placeholder: 'CM-LanguageComprehension',
          // }),
          PropertyPaneTextField('queryModifierProperties.cityMP', {
            label: 'City Managed Property',
            description: 'The managed property name for City', 
            placeholder: 'CM-CityId',
          }),
          PropertyPaneTextField('queryModifierProperties.deadlineFilterMP', {
            label: 'ApplicationDeadlineDate Filter Managed Property',
            description: 'The filter managed property name for ApplicationDeadlineDate', 
            placeholder: 'RefinableDateFirst00',
          }),
          PropertyPaneTextField('queryModifierProperties.jobTypeMP', {
            label: 'JobType Managed Property',
            description: 'The property name for JobType', 
            placeholder: 'CM-JobType',
          }),
          PropertyPaneTextField('queryModifierProperties.workArrangementMP', {
            label: 'WorkArrangement Managed Property',
            description: 'The managed property name for WorkArrangement', 
            placeholder: 'CM-WorkArrangementId',
          })
        ],
      },
    ];
  }
}