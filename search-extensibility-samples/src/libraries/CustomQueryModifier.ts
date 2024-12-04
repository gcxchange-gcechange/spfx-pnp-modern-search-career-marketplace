import { BaseQueryModifier } from "@pnp/modern-search-extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as myLibraryStrings from 'MyCompanyLibraryLibraryStrings';
import { Globals, Language } from "./Globals";

export interface IAdvancedSearchQueryModifierProperties {
  searchBoxSelector: string;
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

  private jobTitle: string = '*';
  private department: string = '*';
  private classificationCode: string = '*';
  private classificationLevel: string = '*';
  private languageRequirement: string = '*';
  private region: string = '*';
  private duration: string = '*';

  // TODO: Search when prop change
  public async onInit(): Promise<void> {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const context = this;

    if (this._properties.jobTitleId) {
      let el = document.getElementById(this._properties.jobTitleId);
      if (el) {
        el.addEventListener('change', function(event) {
          let target = event.target as HTMLInputElement;
          context.jobTitle = target.value;
          context.triggerSearch();
        });
      }
    }

    if (this._properties.departmentId) {
      let el = document.getElementById(this._properties.departmentId);
      if (el) {
        el.addEventListener('focusin', (event) => {
          context.department = context.getElementText(context._properties.departmentId);
          if (context.department != '*')
            context.triggerSearch();
        });
      }
    }

    if (this._properties.classificationCodeId) {
      let el = document.getElementById(this._properties.classificationCodeId);
      if (el) {
        el.addEventListener('focusin', (event) => {
          context.classificationCode = context.getElementText(context._properties.classificationCodeId);
          if (context.classificationCode != '*')
            context.triggerSearch();
        });
      }
    }

    if (this._properties.classificationLevelId) {
      let el = document.getElementById(this._properties.classificationLevelId);
      if (el) {
        el.addEventListener('focusin', (event) => {
          context.classificationLevel = context.getElementText(context._properties.classificationLevelId);
          if (context.classificationLevel != '*')
            context.triggerSearch();
        });
      }
    }

    if (this._properties.languageRequirementId) {
      let el = document.getElementById(this._properties.languageRequirementId);
      if (el) {
        el.addEventListener('focusin', (event) => {
          context.languageRequirement = context.getElementText(context._properties.languageRequirementId);
          if (context.languageRequirement != '*')
            context.triggerSearch();
        });
      }
    }

    if (this._properties.regionId) {
      let el = document.getElementById(this._properties.regionId);
      if (el) {
        el.addEventListener('focusin', (event) => {
          context.region = context.getElementText(context._properties.regionId);
          if (context.region != '*')
            context.triggerSearch();
        });
      }
    }
    
    if (this._properties.durationId) {
      let el = document.getElementById(this._properties.durationId);
      if (el) {
        el.addEventListener('focusin', (event) => {
          context.duration = context.getElementText(context._properties.durationId);
          if (context.duration != '*')
            context.triggerSearch();
        });
      }
    }
  }

  private getElementText(id: string): string | null {
    let el = document.getElementById(id);
    if (el) {
      let retVal = el.innerText;

      if (el.nodeName === 'INPUT')
        retVal = (el as HTMLTextAreaElement).defaultValue;

      if (retVal == '' || retVal == '')
        retVal = '*';

      return retVal.replace('\n', '');
    }
    return '*';
  }

  private triggerSearch() {
    if (this._properties.searchBoxSelector) {
      let el = document.querySelector(this._properties.searchBoxSelector);
      if (el) {
        let searchBox = el as HTMLInputElement;

        if (searchBox.defaultValue == "") {
          searchBox.value = " ";
          searchBox.dispatchEvent(new Event('input', { bubbles: true }));
        }

        el.dispatchEvent(new KeyboardEvent('keydown', {
          key: 'Enter',
          code: 'Enter',
          keyCode: 13,
          which: 13,
          bubbles: true,
          cancelable: true 
        }));
      }
    }
  }

  public async modifyQuery(queryText: string): Promise<string> {
    queryText = queryText === undefined ? '*' : queryText;

    this.jobTitle = this._properties.jobTitleId ? this.getElementText(this._properties.jobTitleId) : '*';
    this.department = this._properties.departmentId ? this.getElementText(this._properties.departmentId) : '*';
    this.classificationCode = this._properties.classificationCodeId ? this.getElementText(this._properties.classificationCodeId) : '*';
    this.classificationLevel = this._properties.classificationLevelId ? this.getElementText(this._properties.classificationLevelId) : '*';
    this.languageRequirement = this._properties.languageRequirementId ? this.getElementText(this._properties.languageRequirementId) : '*';
    this.region = this._properties.regionId ? this.getElementText(this._properties.regionId) : '*';
    this.duration = this._properties.durationId ? this.getElementText(this._properties.durationId) : '*';

    if (Globals.getLanguage() == Language.French) {
      return `${queryText} path: https://devgcx.sharepoint.com/sites/CM-test/Lists/JobOpportunity/ contentclass: STS_ListItem_GenericList "CM-JobTitleFr":*${this.jobTitle}* AND "CM-LanguageRequirement":${this.languageRequirement} AND "CM-Department":${this.department} AND "CM-ClassificationCode":${this.classificationCode} AND "CM-ClassificationLevel":${this.classificationLevel} AND "CM-Duration":${this.duration}`;
    }
    else 
      return `${queryText} path: https://devgcx.sharepoint.com/sites/CM-test/Lists/JobOpportunity/ contentclass: STS_ListItem_GenericList "CM-JobTitleEn":*${this.jobTitle}* AND "CM-LanguageRequirement":${this.languageRequirement} AND "CM-Department":${this.department} AND "CM-ClassificationCode":${this.classificationCode} AND "CM-ClassificationLevel":${this.classificationLevel} AND "CM-Duration":${this.duration}`;
  }

  // TODO: Update listeners
  public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any): void {
    switch(propertyPath) {
      case 'queryModifierProperties.searchBoxSelector':
        break;
      case 'queryModifierProperties.jobTitleId':
        break;
      case 'queryModifierProperties.departmentId':
        break;
      case 'queryModifierProperties.classificationCodeId':
        break;
      case 'queryModifierProperties.classificationLevelId':
        break;
      case 'queryModifierProperties.languageRequirementId':
        break;
      case 'queryModifierProperties.regionId':
        break;
      case 'queryModifierProperties.durationId':
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
          PropertyPaneTextField('queryModifierProperties.jobTitleId', {
            label: 'JobTitle ID',
            description: 'The ID of the HTML element containing the job title text box.',
            placeholder: 'txtJobTitle',
          }),
          PropertyPaneTextField('queryModifierProperties.departmentId', {
            label: 'Department ID',
            description: 'The ID of the HTML element containing the department drop down.',
            placeholder: 'ddDepartment',
          }),
          PropertyPaneTextField('queryModifierProperties.classificationCodeId', {
            label: 'ClassificationCode ID',
            description: 'The ID of the HTML element containing the classification code drop down.',
            placeholder: 'ddClassificationCode',
          }),
          PropertyPaneTextField('queryModifierProperties.classificationLevelId', {
            label: 'ClassificationLevel ID',
            description: 'The ID of the HTML element containing the classification level drop down.',
            placeholder: 'ddClassificationLevel',
          }),
          PropertyPaneTextField('queryModifierProperties.languageRequirementId', {
            label: 'LanguageRequirement ID',
            description: 'The ID of the HTML element containing the language requirement drop down.',
            placeholder: 'ddLanguageRequirement',
          }),
          PropertyPaneTextField('queryModifierProperties.regionId', {
            label: 'Region ID',
            description: 'The ID of the HTML element containing the region drop down.',
            placeholder: 'ddRegion',
          }),
          PropertyPaneTextField('queryModifierProperties.durationId', {
            label: 'Duration ID',
            description: 'The ID of the HTML element containing the duration drop down.',
            placeholder: 'ddDuration',
          })
        ],
      },
    ];
  }
}