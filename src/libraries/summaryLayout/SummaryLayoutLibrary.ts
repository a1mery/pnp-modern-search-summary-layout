import {
  IAdaptiveCardAction,
  IComponentDefinition,
  IDataSourceDefinition,
  IExtensibilityLibrary,
  ILayoutDefinition,
  IQueryModifierDefinition,
  ISuggestionProviderDefinition,
  LayoutType,
  LayoutRenderType,
  BaseLayout
} from '@pnp/modern-search-extensibility';
import { DocumentSummaryWebComponent, DEFAULT_COPILOT_PROMPT } from './components/SummaryComponent';
import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { IPropertyPaneField, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'SummaryLayoutLibraryStrings';

export interface ISummaryLayoutProperties {
  copilotPromptText: string;
}

export class SummaryLayout extends BaseLayout<ISummaryLayoutProperties> {

  public onInit(): void {
    this.properties.copilotPromptText = this.properties.copilotPromptText || DEFAULT_COPILOT_PROMPT;
  }

  public getPropertyPaneFieldsConfiguration(_availableFields: string[]): IPropertyPaneField<unknown>[] {
    return [
      PropertyPaneTextField('copilotPromptText', {
        label: strings.CopilotPromptFieldLabel,
        description: strings.CopilotPromptFieldDescription,
        multiline: true,
        rows: 4,
        value: this.properties.copilotPromptText
      })
    ];
  }
}

export class SummaryLayoutLibrary implements IExtensibilityLibrary {

  public static readonly serviceKey: ServiceKey<SummaryLayoutLibrary> =
    ServiceKey.create<SummaryLayoutLibrary>('SPFx:SummaryLayoutLibrary', SummaryLayoutLibrary);

  constructor(_serviceScope: ServiceScope) {
    // Constructor intentionally left minimal
  }

  public getCustomLayouts(): ILayoutDefinition[] {
    return [
      {
        name: 'Document Summary',
        iconName: 'TextDocument',
        key: 'DocumentSummaryLayout',
        type: LayoutType.Results,
        renderType: LayoutRenderType.Handlebars,
        templateContent: require('./summary-layout.html').default.toString(),
        serviceKey: ServiceKey.create<SummaryLayout>('SummaryLayout', SummaryLayout)
      }
    ];
  }

  public getCustomWebComponents(): IComponentDefinition<unknown>[] {
    return [
      {
        componentName: 'document-summary',
        componentClass: DocumentSummaryWebComponent
      }
    ];
  }

  public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [];
  }

  public registerHandlebarsCustomizations?(handlebarsNamespace: typeof Handlebars): void {
    handlebarsNamespace.registerHelper('getFileTypeAbbreviation', (fileType: string) => {
      if (!fileType) return 'F';
      const abbreviations: Record<string, string> = {
        'docx': 'W',
        'doc': 'W',
        'xlsx': 'X',
        'xls': 'X',
        'pptx': 'P',
        'ppt': 'P',
        'pdf': 'PDF',
        'one': 'N',
        'vsdx': 'V'
      };
      return abbreviations[fileType.toLowerCase()] || fileType.substring(0, 3).toUpperCase();
    });
  }

  public invokeCardAction(_action: IAdaptiveCardAction): void {
    // Not used for Handlebars layout
  }

  public getCustomQueryModifiers?(): IQueryModifierDefinition[] {
    return [];
  }

  public getCustomDataSources?(): IDataSourceDefinition[] {
    return [];
  }

  public name(): string {
    return 'SummaryLayoutLibrary';
  }
}
