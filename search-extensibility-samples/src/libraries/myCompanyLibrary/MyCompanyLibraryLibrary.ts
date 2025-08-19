import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
//import { DynamicProperty } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {  IExtensibilityLibrary, 
          IComponentDefinition, 
          ISuggestionProviderDefinition, 
          ISuggestionProvider,
          ILayoutDefinition, 
          LayoutType, 
          ILayout,
          LayoutRenderType,
          IQueryModifierDefinition,
          IQueryModifier,
          IDataSourceDefinition,
          IDataSource
} from "@pnp/modern-search-extensibility";
import * as Handlebars from "handlebars";
import { MyCustomComponentWebComponent } from "../CustomComponent";
import { CustomLayout } from "../CustomLayout";
import { CustomSuggestionProvider } from "../CustomSuggestionProvider";
import { AdvancedSearchQueryModifier } from "../CustomQueryModifier";
import { CustomDataSource } from "../CustomDataSource";
import { SelectLanguage } from "../SelectLanguage";
import { Globals } from "../Globals";
import { MyOpportunitiesQueryModifier } from "../MyOpportunitiesQueryModifier";

export class MyCompanyLibraryLibrary implements IExtensibilityLibrary {
  

  public static readonly serviceKey: ServiceKey<MyCompanyLibraryLibrary> =
  ServiceKey.create<MyCompanyLibraryLibrary>('SPFx:MyCustomLibraryComponent', MyCompanyLibraryLibrary);

  private _spHttpClient: SPHttpClient;
  private _pageContext: PageContext;
  private _currentWebUrl: string;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
      this._pageContext = serviceScope.consume(PageContext.serviceKey);
      this._currentWebUrl = this._pageContext.web.absoluteUrl;

      Globals.userDisplayName = this._pageContext.user.displayName;
    });
  }

  public getCustomLayouts(): ILayoutDefinition[] {
    return [
      {
        name: 'Job Opportunity',
        iconName: 'Suitcase',
        key: 'CustomLayoutHandlebars',
        type: LayoutType.Results,
        renderType: LayoutRenderType.Handlebars,
        templateContent: require('../JobOpportunity.results.html'),
        serviceKey: ServiceKey.create<ILayout>('PnP:CustomLayoutHandlebars', CustomLayout),
      },
      {
        name: 'PnP Custom layout (Adaptive Cards)',
        iconName: 'Color',
        key: 'CustomLayoutAdaptive',
        type: LayoutType.Results,
        renderType: LayoutRenderType.AdaptiveCards,
        templateContent: JSON.stringify(require('../custom-layout.json'), null, "\t"),
        serviceKey: ServiceKey.create<ILayout>('PnP:CustomLayoutAdaptive', CustomLayout),
      }
    ];
  }

  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [
      {
        componentName: 'job-opportunity-card',
        componentClass: MyCustomComponentWebComponent
      }
    ];
  }

  public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [
        {
          name: 'Custom Suggestions Provider',
          key: 'CustomSuggestionsProvider',
          description: 'A demo custom suggestions provider from the extensibility library',
          serviceKey: ServiceKey.create<ISuggestionProvider>('MyCompany:CustomSuggestionsProvider', CustomSuggestionProvider)
      }
    ];
  }

  public registerHandlebarsCustomizations(namespace: typeof Handlebars) {
    namespace.registerHelper('results', (value: any) => {
      try {
        if (value['string'].indexOf('\'') != -1)
          return new namespace.SafeString(
            `${value['string'].replace('results for', SelectLanguage(Globals.getLanguage()).resultsFor)}`
          );
        else 
          return new namespace.SafeString(
            `${value['string'].replace('results', SelectLanguage(Globals.getLanguage()).results)}`
          );
      }
      catch (e) {
        console.log(e);
        return value;
      }
    });

    namespace.registerHelper('resultsNoQueryText', (value: any) => {
      try {
        return new namespace.SafeString(
          `${value['string']
            .replace(' for ', '')
            .replace(' de ', '')
            .replace('« ','')
            .replace(' »','')
            .replace('<em>[object Object]</em>', '')
            .replace('\'<em>[object Object]</em>\'', '')
            .replace('\'\'','')
            .replace('results', SelectLanguage(Globals.getLanguage()).results)}`
        );
      }
      catch (e) {
        console.log(e);
        return value;
      }
    });

    namespace.registerHelper('term', (value: string) => {
      try {
        const matches = value.match(/L0\|#[^|]+\|(.+)/);
        return matches && matches[1] ? matches[1].trim() : value;
      }
      catch (e) {
        console.log(e);
        return value;
      }
    });

    namespace.registerHelper('terms', (value: string) => {
      try {
        if (value){
          let terms = [];
          let split = value.split(';GTSet');
          for (let i = 0; i < split.length - (split.length > 1 ? 1 : 0); i++) {
            const parts = split[i].split('|');
            terms.push(parts[parts.length - 1]);
          }
          return terms.join(', ');
        }
        return value;
      }
      catch (e) {
        console.log(e);
        return value;
      }
    });

    namespace.registerHelper('opportunitiesLabel', () => {
      try {
        const strings = SelectLanguage(Globals.getLanguage());
        return strings.opportunities;
      }
      catch (e) {
        console.log(e);
        return '';
      }
    });
  }

  public invokeCardAction(action: any): void {
    
    // Process the action based on type
    if (action.type == "Action.OpenUrl") {
      window.open(action.url, "_blank");
    } else if (action.type == "Action.Submit") {
      // Process the action based on title
      switch (action.title) {

        case 'Click on item':

           // Invoke the currentUser endpoing
           this._spHttpClient.get(
            `${this._currentWebUrl}/_api/web/currentUser`,
            SPHttpClient.configurations.v1, 
            null).then((response: SPHttpClientResponse) => {
              response.json().then((json) => {
                console.log(JSON.stringify(json));
              });
            });

          break;

        case 'Global click':
          alert(JSON.stringify(action.data));
          break;
        default:
          console.log('Action not supported!');
          break;
      }
    }
  }

  public getCustomQueryModifiers(): IQueryModifierDefinition[]
  {
    return [
      {
        name: 'Advanced Search',
        key: 'AdvancedSearch',
        description: 'A query modifier for career marketplace advanced search.',
        serviceKey: ServiceKey.create<IQueryModifier>('MyCompany:CustomQueryModifier', AdvancedSearchQueryModifier)

      },
      {
        name: 'Owner Opportunities',
        key: 'MyOpportunities',
        description: 'A query modifier for career marketplace which returns the logged in user\'s job opportunities that they have posted to the system.',
        serviceKey: ServiceKey.create<IQueryModifier>('MyCompany:MyOpportunitiesQueryModifier', MyOpportunitiesQueryModifier)
      }
    ];
  }

  public getCustomDataSources(): IDataSourceDefinition[] {
    return [
      {
          name: 'NPM Search',
          iconName: 'Database',
          key: 'CustomDataSource',
          serviceKey: ServiceKey.create<IDataSource>('MyCompany:CustomDataSource', CustomDataSource)
      }
    ];
  }

  public name(): string {
    return 'CareerMarketplaceLibraryComponent';
  }
}
