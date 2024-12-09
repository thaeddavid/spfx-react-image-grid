import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldFilePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';

import * as strings from 'ShowcaseGridWebPartStrings';
import ShowcaseGrid from './components/ShowcaseGrid';
import { IShowcaseItem, IShowcaseGridProps } from './components/IShowcaseGridProps';

export interface IShowcaseGridWebPartProps {
  gridItems: IShowcaseItem[];
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
        gridItems: this.properties.gridItems || []
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
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "Grid Item 1",
              groupFields: [
                PropertyFieldFilePicker('gridItems[0].imageUrl', {
                  context: this.context as any,
                  filePickerResult: {
                    fileAbsoluteUrl: this.properties.gridItems?.[0]?.imageUrl || '',
                    fileName: '',
                    fileNameWithoutExtension: '',
                    downloadFileContent: () => Promise.resolve('')
                  } as any,
                  onPropertyChange: (propertyPath: string, newValue: any) => {
                    console.log('File picker value:', newValue);
                    if (newValue && typeof newValue === 'object') {
                      this.properties.gridItems = this.properties.gridItems || [];
                      this.properties.gridItems[0] = this.properties.gridItems[0] || {};
                      this.properties.gridItems[0].imageUrl = 
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
                  key: "filePickerFieldId1"
                }),
                PropertyPaneTextField('gridItems[0].title', { label: "Title" }),
                PropertyPaneTextField('gridItems[0].description', { 
                  label: "Description",
                  multiline: true,
                  rows: 6,
                  resizable: true 
                }),
                PropertyPaneTextField('gridItems[0].linkUrl', { label: "Link URL" }),
                PropertyPaneTextField('gridItems[0].linkText', { label: "Link Text" })
              ]
            },
            {
              groupName: "Grid Item 2",
              groupFields: [
                PropertyFieldFilePicker('gridItems[1].imageUrl', {
                  context: this.context as any,
                  filePickerResult: {
                    fileAbsoluteUrl: this.properties.gridItems?.[1]?.imageUrl || '',
                    fileName: '',
                    fileNameWithoutExtension: '',
                    downloadFileContent: () => Promise.resolve('')
                  } as any,
                  onPropertyChange: (propertyPath: string, newValue: any) => {
                    console.log('File picker value:', newValue);
                    if (newValue && typeof newValue === 'object') {
                      this.properties.gridItems = this.properties.gridItems || [];
                      this.properties.gridItems[1] = this.properties.gridItems[1] || {};
                      this.properties.gridItems[1].imageUrl = 
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
                  key: "filePickerFieldId2"
                }),
                PropertyPaneTextField('gridItems[1].title', { label: "Title" }),
                PropertyPaneTextField('gridItems[1].description', { 
                  label: "Description",
                  multiline: true,
                  rows: 6,
                  resizable: true 
                }),
                PropertyPaneTextField('gridItems[1].linkUrl', { label: "Link URL" }),
                PropertyPaneTextField('gridItems[1].linkText', { label: "Link Text" })
              ]
            },
            {
              groupName: "Grid Item 3",
              groupFields: [
                PropertyFieldFilePicker('gridItems[2].imageUrl', {
                  context: this.context as any,
                  filePickerResult: {
                    fileAbsoluteUrl: this.properties.gridItems?.[2]?.imageUrl || '',
                    fileName: '',
                    fileNameWithoutExtension: '',
                    downloadFileContent: () => Promise.resolve('')
                  } as any,
                  onPropertyChange: (propertyPath: string, newValue: any) => {
                    console.log('File picker value:', newValue);
                    if (newValue && typeof newValue === 'object') {
                      this.properties.gridItems = this.properties.gridItems || [];
                      this.properties.gridItems[2] = this.properties.gridItems[2] || {};
                      this.properties.gridItems[2].imageUrl = 
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
                  key: "filePickerFieldId3"
                }),
                PropertyPaneTextField('gridItems[2].title', { label: "Title" }),
                PropertyPaneTextField('gridItems[2].description', { 
                  label: "Description",
                  multiline: true,
                  rows: 6,
                  resizable: true 
                }),
                PropertyPaneTextField('gridItems[2].linkUrl', { label: "Link URL" }),
                PropertyPaneTextField('gridItems[2].linkText', { label: "Link Text" })
              ]
            },
            {
              groupName: "Grid Item 4",
              groupFields: [
                PropertyFieldFilePicker('gridItems[3].imageUrl', {
                  context: this.context as any,
                  filePickerResult: {
                    fileAbsoluteUrl: this.properties.gridItems?.[3]?.imageUrl || '',
                    fileName: '',
                    fileNameWithoutExtension: '',
                    downloadFileContent: () => Promise.resolve('')
                  } as any,
                  onPropertyChange: (propertyPath: string, newValue: any) => {
                    console.log('File picker value:', newValue);
                    if (newValue && typeof newValue === 'object') {
                      this.properties.gridItems = this.properties.gridItems || [];
                      this.properties.gridItems[3] = this.properties.gridItems[3] || {};
                      this.properties.gridItems[3].imageUrl = 
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
                  key: "filePickerFieldId4"
                }),
                PropertyPaneTextField('gridItems[3].title', { label: "Title" }),
                PropertyPaneTextField('gridItems[3].description', { 
                  label: "Description",
                  multiline: true,
                  rows: 6,
                  resizable: true 
                }),
                PropertyPaneTextField('gridItems[3].linkUrl', { label: "Link URL" }),
                PropertyPaneTextField('gridItems[3].linkText', { label: "Link Text" })
              ]
            }
          ]
        }
      ]
    };
  }
}
