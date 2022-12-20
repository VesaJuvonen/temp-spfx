import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { HelloWorldPropertyPane } from './HelloWorldPropertyPane';
import { SPHttpClient } from '@microsoft/sp-http';

export interface IHelloWorldAdaptiveCardExtensionProps {
  title: string;
  description: string;
  iconProperty: string;
  listId: string;
}

export interface IHelloWorldAdaptiveCardExtensionState {
  currentIndex: number;
  items: IListItem[];
}

export interface IListItem {
  title: string;
  description: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'HelloWorld_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'HelloWorld_QUICK_VIEW';

export default class HelloWorldAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IHelloWorldAdaptiveCardExtensionProps,
  IHelloWorldAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: HelloWorldPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = {
      currentIndex: 0,
      items: []
     };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return this._fetchData();
  }

  private _fetchData(): Promise<void> {
    if (this.properties.listId) {
      return this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}` +
          `/_api/web/lists/GetById(id='${this.properties.listId}')/items`,
        SPHttpClient.configurations.v1
      )
        .then((response) => response.json())
        .then((jsonResponse) => {
          if (!jsonResponse.value) {
            return Promise.reject("Could not find the list")
          }
          else
          {
            return jsonResponse.value.map(
              (item: { Title: any; Description: any; }) => { return { title: item.Title, description: item.Description }; })
          }
        })  
        .then((items) => this.setState({ items }));
    }
  
    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'HelloWorld-property-pane'*/
      './HelloWorldPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.HelloWorldPropertyPane();
        }
      );
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'listId' && newValue !== oldValue) {
      if (newValue) {
        this._fetchData();
      } else {
        this.setState({ items: [] });
      }
    }
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }
}
