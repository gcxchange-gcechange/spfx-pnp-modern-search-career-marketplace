import { BaseQueryModifier } from "@pnp/modern-search-extensibility";

export interface ITestQueryModifierProperties {
  test: string;
}

export class TestQueryModifier extends BaseQueryModifier<ITestQueryModifierProperties> {
  
  public async onInit(): Promise<void> {

  }

  public async modifyQuery(queryText: string): Promise<string> {
    queryText = "Test";

    console.log(queryText);

    return queryText;
  }
}