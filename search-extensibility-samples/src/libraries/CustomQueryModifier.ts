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
  //languageComprehensionMP: string;
  cityMP: string;
  durationMP: string;
  durationQuantityMP: string;
  deadlineFilterMP: string;
  jobTypeMP: string;
  programAreaMP: string;
  durationYearsId: string;
  durationMonthsId: string;
  durationWeeksId: string;
}

enum AdvancedSearchSessionKeys {
  Initialized = 'gcx-cm-adsearch-init',
  JobTitle = 'gcx-cm-adsearch-jobTitle',
  ClassificationCode = 'gcx-cm-adsearch-classificationCode',
  ClassificationLevel = 'gcx-cm-adsearch-classificationLevel',
  Department = 'gcx-cm-adsearch-department',
  Duration = 'gcx-cm-adsearch-duration',
  DurationQuantity = 'gcx-cm-adsearch-durationQuantity',
  DurationOperator = 'gcx-cm-adsearch-durationOperator',
  LanguageRequirement = 'gcx-cm-adsearch-languageRequirement',
  //LanguageComprehension = 'gcx-cm-adsearch-languageComprehension',
  City = 'gcx-cm-adsearch-city'
}

enum FilterSessionKeys {
  Initialized = 'gcx-cm-filter-init',
  JobType = 'gcx-cm-filter-jobType',
  ProgramArea = 'gcx-cm-filter-programArea',
  ApplicationDeadline = 'gcx-cm-filter-applicationDeadline',
}

export enum QueryModifierKeys {
  AdvancedSearch = 'gcx-cm-querymod-as',
  MyOpportunities = 'gcx-cm-querymod-myop'
}

//CustomQueryModifier
export class AdvancedSearchQueryModifier extends BaseQueryModifier<IAdvancedSearchQueryModifierProperties> {
  private static readonly DEFAULT_VALUE = '*';
  private lang = Globals.getLanguage();

  public async onInit(): Promise<void> {

    // Initialize the advanced search session storage items
    (Object.keys(AdvancedSearchSessionKeys) as (keyof typeof AdvancedSearchSessionKeys)[]).forEach(key => {
      const value = AdvancedSearchSessionKeys[key];
      if (!sessionStorage.getItem(value) && value !== AdvancedSearchSessionKeys.Initialized)
        sessionStorage.setItem(value, '');
    });

    // Initialize the filter session storage items
    (Object.keys(FilterSessionKeys) as (keyof typeof FilterSessionKeys)[]).forEach(key => {
      const value = FilterSessionKeys[key];
      if (value !== FilterSessionKeys.Initialized)
        sessionStorage.setItem(value, '');
    });

    this.setupListeners();
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

    tryGetSessionStorageItem(AdvancedSearchSessionKeys.Initialized, () => {
      // Advanced Search - Search Btn
      tryGetElement(this._properties.searchButtonId, (el: HTMLElement) => {
        el.addEventListener('click', (event) => {
          event.preventDefault();
          setTimeout(() => {
            context.triggerSearch();
          }, 0);
        });
      });
      // Advanced Search - Clear Btn
      tryGetElement(this._properties.clearButtonId, (el: HTMLElement) => {
        el.addEventListener('click', (event) => {
          event.preventDefault();
          setTimeout(() => {
            // Clear the pnp search box & retrigger search
            context.triggerSearch(true);
          }, 0);
        });
      });
    });

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

    let finalQuery = this.applyAdvancedSearch(`${queryText !== '*' ? '*' + queryText + '*' : queryText} path: ${this._properties.listPath} contentclass: STS_ListItem_GenericList`);
    finalQuery = this.applyFilters(finalQuery);

    // Set this item so the other custom queries know we've already performed an advanced search/filter on the original query
    sessionStorage.setItem(QueryModifierKeys.AdvancedSearch, 'true');

    return finalQuery;
  }

  private applyAdvancedSearch(query: string): string {
    let finalQuery = `${query} `;
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

      // const languageComprehension = sessionStorage.getItem(AdvancedSearchSessionKeys.LanguageComprehension);
      // if (languageComprehension && languageComprehension.trim() != '') {
      //   finalQuery += ` AND ("${this._properties.languageComprehensionMP}":`;
      //   let comprehensions = this.getAllLanguageComprehensions(languageComprehension);
      //   for (let i = 0; i < comprehensions.length; i++) {
      //     finalQuery += `${i > 0 ? 'OR ' : ''}"${comprehensions[i]}" `;
      //   }
      //   finalQuery += ')';
      // }
    }

    const duration = sessionStorage.getItem(AdvancedSearchSessionKeys.Duration);
    const durationQuantity = sessionStorage.getItem(AdvancedSearchSessionKeys.DurationQuantity);
    const durationOperator = sessionStorage.getItem(AdvancedSearchSessionKeys.DurationOperator);
    if (duration && duration.trim() != '' &&
        durationQuantity && durationQuantity.trim() != '' &&
        durationOperator && durationOperator.trim() != '')  {
      try {
        const operator = durationOperator === '0' ? '=' : (durationOperator === '2' ? '<=' : '>=');
        let durationInDays: number;

        if (this._properties.durationYearsId && this._properties.durationYearsId != '' && 
          this._properties.durationMonthsId && this._properties.durationMonthsId != '' &&
          this._properties.durationWeeksId && this._properties.durationWeeksId != '') {
            
          switch (duration) {
          case this._properties.durationYearsId:
              durationInDays = 365 * parseInt(durationQuantity);
              break;
            case this._properties.durationMonthsId:
              durationInDays = Math.round(365 / 12 * parseInt(durationQuantity));
              break;
            case this._properties.durationWeeksId:
              durationInDays = Math.round(365 / 52 * parseInt(durationQuantity));
              break;
            default:
              throw new Error(`Couldn't map Duration:${duration} to any of the following: [durationYearsId:${this._properties.durationYearsId}, durationMonthsId:${this._properties.durationMonthsId}, durationWeeksId:${this._properties.durationWeeksId}]`);
          }
        }
        else {
          console.error(`One of the following is not configured: [durationYearsId:${this._properties.durationYearsId}, durationMonthsId:${this._properties.durationMonthsId}, durationWeeksId:${this._properties.durationWeeksId}]`);
        }

        // If we're not searching for an exact duration include results for Deployment job types (they always have 0 DurationInDays)
        if (operator !== '=') {
          finalQuery += `${propSet ? 'AND (' : '('}${this._properties.durationQuantityMP}${operator}${durationInDays} OR ${this._properties.durationQuantityMP}=0) `;
        }
        else {
          finalQuery += `${propSet ? 'AND ' : ''}${this._properties.durationQuantityMP}${operator}${durationInDays} `;
        }

        propSet = true;
      } catch(e) {
        console.error('Couldn\'t performance advanced search on Duration because an error occured.');
        console.error(e);
      }
    }

    const city = sessionStorage.getItem(AdvancedSearchSessionKeys.City);
    if (city && city.trim() != '') {
      finalQuery += `${propSet ? 'AND ' : ''}"${this._properties.cityMP}":${city} `;
      propSet = true;
    }

    // Only show results where the ApplicationDeadlineDate is today's date or greater
    const today = new Date();
    const formattedUTCDate = `${today.getUTCMonth() + 1}/${today.getUTCDate()}/${today.getUTCFullYear()}`;

    finalQuery += `AND "${this._properties.deadlineFilterMP}">=${formattedUTCDate}`;

    return finalQuery;
  }

  private applyFilters(query: string): string {
    const applicationDeadline = sessionStorage.getItem(FilterSessionKeys.ApplicationDeadline);
    const jobTypes = sessionStorage.getItem(FilterSessionKeys.JobType);
    const programAreas = sessionStorage.getItem(FilterSessionKeys.ProgramArea);

    if ((applicationDeadline === undefined || applicationDeadline.trim() == '')
     && (jobTypes === undefined || jobTypes.trim() == '')
     && (programAreas === undefined || programAreas.trim() == '')) {
      return query;
    }

    let finalQuery = `${query} AND (`;

    if (applicationDeadline && applicationDeadline.trim() != '') {
      finalQuery += `"${this._properties.deadlineFilterMP}"<=${applicationDeadline} `;
    } else {
      const today = new Date();
      const formattedUTCDate = `${today.getUTCMonth() + 1}/${today.getUTCDate()}/${today.getUTCFullYear()}`;
      finalQuery += `"${this._properties.deadlineFilterMP}">=${formattedUTCDate} `;
    }

    if (jobTypes && jobTypes.trim() != '') {
      const jobTypeArr = jobTypes.split(',');
      for (let i = 0; i < jobTypeArr.length; i++) {

        if (i == 0)
          finalQuery += `AND (`;

        finalQuery += `"${this._properties.jobTypeMP}":${jobTypeArr[i]}`;

        if (i != jobTypeArr.length - 1)
          finalQuery += ' OR ';
      }

      finalQuery += ')';
    }

    if(programAreas && programAreas.trim() != '') {
      const programAreaArr = programAreas.split(',');
      for (let i = 0; i < programAreaArr.length; i++) {

        if (i == 0)
          finalQuery += `AND (`;

        finalQuery += `"${this._properties.programAreaMP}":${programAreaArr[i]}`;

        if (i != programAreaArr.length - 1)
          finalQuery += ' OR ';
      }

      finalQuery += ')';
    }

    return `${finalQuery})`;
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
          PropertyPaneTextField('queryModifierProperties.durationMP', {
            label: 'Duration Managed Property',
            description: 'The managed property name for Duration', 
            placeholder: 'CM-DurationId',
          }),
          PropertyPaneTextField('queryModifierProperties.durationQuantityMP', {
            label: 'Duration Quantity Managed Property',
            description: 'The managed property name for DurationQuantity', 
            placeholder: 'RefinableInt00',
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
          PropertyPaneTextField('queryModifierProperties.programAreaMP', {
            label: 'ProgramArea Managed Property',
            description: 'The managed property name for ProgramArea', 
            placeholder: 'CM-ProgramArea',
          }),
          PropertyPaneTextField('queryModifierProperties.durationYearsId', {
            label: 'Duration Years ID',
            description: 'The ID in the Duration list for the "year(s)" list item'
          }),
          PropertyPaneTextField('queryModifierProperties.durationMonthsId', {
            label: 'Duration Months ID',
            description: 'The ID in the Duration list for the "month(s)" list item'
          }),
          PropertyPaneTextField('queryModifierProperties.durationWeeksId', {
            label: 'Duration Weeks ID',
            description: 'The ID in the Duration list for the "week(s)" list item'
          })
        ],
      },
    ];
  }
}