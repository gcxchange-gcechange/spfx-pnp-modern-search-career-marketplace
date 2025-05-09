import { BaseQueryModifier } from "@pnp/modern-search-extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as myLibraryStrings from 'MyCompanyLibraryLibraryStrings';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs"; 
import "@pnp/sp/site-users";
import { QueryModifierKeys } from "./CustomQueryModifier";

export interface IMyOpportunitiesQueryModifierProperties {
  listPath: string;
  contactObjectIdMP: string;
}

//CustomQueryModifier
export class MyOpportunitiesQueryModifier extends BaseQueryModifier<IMyOpportunitiesQueryModifierProperties> {
    private _userId: string;
    private _sp: SPFI;

  public async onInit(): Promise<void> {
    this._sp = spfi().using(SPFx(this.context));

    const currentUser = await this._sp.web.currentUser();
    this._userId = currentUser.UserId.NameId;
  }

  public async modifyQuery(queryText: string): Promise<string> {
    const advancedSearchQueryEnabled = sessionStorage.getItem(QueryModifierKeys.AdvancedSearch);
    if (advancedSearchQueryEnabled && advancedSearchQueryEnabled === 'true') {
        sessionStorage.removeItem(QueryModifierKeys.AdvancedSearch);

        queryText += ` AND "${this._properties.contactObjectIdMP}":${this._userId}`;

        return queryText;
    }

    queryText = queryText || '*';

    if (queryText.trim() == '')
      queryText = '*';

    let finalQuery = `${queryText !== '*' ? '*' + queryText + '*' : queryText} path: ${this._properties.listPath} contentclass: STS_ListItem_GenericList`;
    finalQuery += ` AND "${this._properties.contactObjectIdMP}":${this._userId}`;

    return finalQuery;
  }

  public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any): void {

  }

  public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {
    return [
      {
        groupName: myLibraryStrings.MyOpportunitiesQueryModifier.GroupName,
        groupFields: [
          PropertyPaneTextField('queryModifierProperties.listPath', {
            label: 'JobOpportunity List Path',
            description: 'The path to the JobOpportunity list on the site this webpart is deployed.',
            placeholder: 'https://devgcx.sharepoint.com/sites/CM-test/Lists/JobOpportunity/',
          }),
          PropertyPaneTextField('queryModifierProperties.contactObjectIdMP', {
            label: 'ContactObjectId Managed Property Name',
            description: 'The managed property name that maps to the ContactObjectId column in the list.',
            placeholder: 'CM-ContactObjectId',
          }),
        ],
      },
    ];
  }
}