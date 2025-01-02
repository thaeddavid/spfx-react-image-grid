import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldFilePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';

import * as strings from 'ShowcaseGridWebPartStrings';
import ShowcaseGrid from './components/ShowcaseGrid';
import { IShowcaseItem, IShowcaseGridProps } from './components/IShowcaseGridProps';

export interface IShowcaseGridWebPartProps {
  gridItems: IShowcaseItem[];
  columns: number;
  rows: number;
}

export default class ShowcaseGridWebPart extends BaseClientSideWebPart<IShowcaseGridWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IShowcaseGridProps> = React.createElement(
      ShowcaseGrid,
      {
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        gridItems: this.properties.gridItems || [],
        columns: this.properties.columns || 2,
        rows: this.properties.rows || 2
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Calculate total number of grid items based on rows and columns
    const totalItems = (this.properties.rows || 2) * (this.properties.columns || 2);
    
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "Grid Layout",
              groupFields: [
                PropertyPaneSlider('columns', {
                  label: "Number of Columns",
                  min: 1,
                  max: 3,
                  value: this.properties.columns || 2,
                  showValue: true,
                  step: 1
                }),
                PropertyPaneSlider('rows', {
                  label: "Number of Rows",
                  min: 1,
                  max: 3,
                  value: this.properties.rows || 2,
                  showValue: true,
                  step: 1
                })
              ]
            },
            ...Array.from({ length: totalItems }, (_, i) => ({
              groupName: `Grid Item ${i + 1}`,
              groupFields: [
                PropertyFieldFilePicker(`gridItems[${i}].imageUrl`, {
                  context: this.context as any,
                  filePickerResult: {
                    fileAbsoluteUrl: this.properties.gridItems?.[i]?.imageUrl || '',
                    fileName: '',
                    fileNameWithoutExtension: '',
                    downloadFileContent: () => Promise.resolve('')
                  } as any,
                  onPropertyChange: (propertyPath: string, newValue: any) => {
                    if (newValue && typeof newValue === 'object') {
                      this.properties.gridItems = this.properties.gridItems || [];
                      this.properties.gridItems[i] = this.properties.gridItems[i] || {};
                      this.properties.gridItems[i].imageUrl = 
                        newValue.serverRelativeUrl || 
                        newValue.fileAbsoluteUrl || 
                        newValue.fileName || 
                        '';
                    }
                  },
                  onSave: (e: any) => {},
                  properties: this.properties,
                  label: "Image",
                  accepts: [".gif", ".jpg", ".jpeg", ".bmp", ".dib", ".tif", ".tiff", ".ico", ".png", ".jxr", ".svg"],
                  buttonLabel: "Choose Image",
                  key: `filePickerFieldId${i + 1}`
                }),
                PropertyPaneTextField(`gridItems[${i}].title`, { label: "Title" }),
                PropertyPaneTextField(`gridItems[${i}].description`, { 
                  label: "Description",
                  multiline: true,
                  rows: 6,
                  resizable: true 
                }),
                PropertyPaneTextField(`gridItems[${i}].linkUrl`, { label: "Link URL" }),
                PropertyPaneTextField(`gridItems[${i}].linkText`, { label: "Link Text" })
              ]
            }))
          ]
        }
      ]
    };
  }
}
