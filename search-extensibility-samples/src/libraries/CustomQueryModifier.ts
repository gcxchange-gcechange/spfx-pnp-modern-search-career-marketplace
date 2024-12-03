import { BaseQueryModifier } from "@pnp/modern-search-extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as myLibraryStrings from 'MyCompanyLibraryLibraryStrings';
import { Globals, Language } from "./Globals";

export interface IAdvancedSearchQueryModifierProperties {
  jobTitleId: string
  departmentId: string
  classificationCodeId: string
  classificationLevelId: string
  languageRequirementId: string
  regionId: string
  durationId: string
}

//CustomQueryModifier
export class AdvancedSearchQueryModifier extends BaseQueryModifier<IAdvancedSearchQueryModifierProperties> {

  private jobTitle: string;
  private department: string;
  private classificationCode: string;
  private classificationLevel: string;
  private languageRequirement: string;
  private region: string;
  private duration: string;

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
        });
      }
    }

    if (this._properties.departmentId) {
      let el = document.getElementById(this._properties.departmentId);
      if (el) {
        el.addEventListener('focusout', (event) => {
          context.department = context.getElementText(context._properties.departmentId);
        });
      }
    }

    if (this._properties.classificationCodeId) {
      let el = document.getElementById(this._properties.classificationCodeId);
      if (el) {
        el.addEventListener('focusout', (event) => {
          context.classificationCode = context.getElementText(context._properties.classificationCodeId);
        });
      }
    }

    if (this._properties.classificationLevelId) {
      let el = document.getElementById(this._properties.classificationLevelId);
      if (el) {
        el.addEventListener('focusout', (event) => {
          context.classificationLevel = context.getElementText(context._properties.classificationLevelId);
        });
      }
    }

    if (this._properties.languageRequirementId) {
      let el = document.getElementById(this._properties.languageRequirementId);
      if (el) {
        el.addEventListener('focusout', (event) => {
          context.languageRequirement = context.getElementText(context._properties.languageRequirementId);
        });
      }
    }

    if (this._properties.regionId) {
      let el = document.getElementById(this._properties.regionId);
      if (el) {
        el.addEventListener('focusout', (event) => {
          context.region = context.getElementText(context._properties.regionId);
        });
      }
    }
    
    if (this._properties.durationId) {
      let el = document.getElementById(this._properties.durationId);
      if (el) {
        el.addEventListener('focusout', (event) => {
          context.duration = context.getElementText(context._properties.durationId);
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

  public async modifyQuery(queryText: string): Promise<string> {
    if (queryText === undefined)
      queryText = '*';

    if (this._properties.jobTitleId)
      this.jobTitle = this.getElementText(this._properties.jobTitleId);

    if (this._properties.departmentId)
      this.department = this.getElementText(this._properties.departmentId);

    if (this._properties.classificationCodeId)
      this.classificationCode = this.getElementText(this._properties.classificationCodeId);
    
    if (this._properties.classificationLevelId) 
      this.classificationLevel = this.getElementText(this._properties.classificationLevelId);
    
    if (this._properties.languageRequirementId) 
      this.languageRequirement = this.getElementText(this._properties.languageRequirementId);
    
    if (this._properties.regionId) 
      this.region = this.getElementText(this._properties.regionId);
    
    if (this._properties.durationId) {
      this.duration = this.getElementText(this._properties.durationId);
    }

    console.log(this);

    // TODO: Fix search by job title for french/english
    if (Globals.getLanguage() == Language.French) {
      return `${queryText} path: https://devgcx.sharepoint.com/sites/CM-test/Lists/JobOpportunity/ contentclass: STS_ListItem_GenericList "CM-JobTitleFr":*${this.jobTitle}* AND "CM-LanguageRequirement":${this.languageRequirement} AND "CM-Department":${this.department} AND "CM-ClassificationCode":${this.classificationCode} AND "CM-ClassificationLevel":${this.classificationLevel} AND "CM-Duration":${this.duration}`;
    }
    else 
      return `${queryText} path: https://devgcx.sharepoint.com/sites/CM-test/Lists/JobOpportunity/ contentclass: STS_ListItem_GenericList "CM-JobTitleEn":*${this.jobTitle}* AND "CM-LanguageRequirement":${this.languageRequirement} AND "CM-Department":${this.department} AND "CM-ClassificationCode":${this.classificationCode} AND "CM-ClassificationLevel":${this.classificationLevel} AND "CM-Duration":${this.duration}`;
  }

  // TODO
  public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any): void {
    switch(propertyPath) {
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
          PropertyPaneTextField('queryModifierProperties.jobTitleId', {
            label: 'JobTitle ID',
            description: 'The ID of the HTML element containing the job title text box.',
            value: 'txtJobTitle',
          }),
          PropertyPaneTextField('queryModifierProperties.departmentId', {
            label: 'Department ID',
            description: 'The ID of the HTML element containing the department drop down.',
            value: 'ddDepartment',
          }),
          PropertyPaneTextField('queryModifierProperties.classificationCodeId', {
            label: 'ClassificationCode ID',
            description: 'The ID of the HTML element containing the classification code drop down.',
            value: 'ddClassificationCode',
          }),
          PropertyPaneTextField('queryModifierProperties.classificationLevelId', {
            label: 'ClassificationLevel ID',
            description: 'The ID of the HTML element containing the classification level drop down.',
            value: 'ddClassificationLevel',
          }),
          PropertyPaneTextField('queryModifierProperties.languageRequirementId', {
            label: 'LanguageRequirement ID',
            description: 'The ID of the HTML element containing the language requirement drop down.',
            value: 'ddLanguageRequirement',
          }),
          PropertyPaneTextField('queryModifierProperties.regionId', {
            label: 'Region ID',
            description: 'The ID of the HTML element containing the region drop down.',
            value: 'ddRegion',
          }),
          PropertyPaneTextField('queryModifierProperties.durationId', {
            label: 'Duration ID',
            description: 'The ID of the HTML element containing the duration drop down.',
            value: 'ddDuration',
          }),
          // PropertyPaneDynamicFieldSet({
          //   label: 'Advanced Search',
          //   fields: [
          //     PropertyPaneDynamicField('queryModifierProperties.advancedSearch', {
          //       label: 'Advanced Search Web Part',
          //     })
          //   ]
          // }),
        ],
      },
    ];
  }
}