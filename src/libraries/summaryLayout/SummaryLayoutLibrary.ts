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
import { DocumentSummaryWebComponent } from './components/SummaryComponent';
import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';

// eslint-disable-next-line @typescript-eslint/no-empty-interface
export interface ISummaryLayoutProperties {
}

export class SummaryLayout extends BaseLayout<ISummaryLayoutProperties> {
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
