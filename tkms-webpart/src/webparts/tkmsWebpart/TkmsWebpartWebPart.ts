import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { PropertyFieldListPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'TkmsWebpartWebPartStrings';
import TkmsWebpart from './components/TkmsWebpart';
import { ITkmsWebpartProps } from './components/ITkmsWebpartProps';

export interface IListViewFieldConfig {
  FieldInternalName: string;
  FieldDisplayName: string;
  FieldName: string;
  FieldType: string;
  FieldConfig?: string;
}

export interface ITkmsWebpartWebPartProps {
  description: string;
  newButtonText?: string;
  listId?: string;
  listViewFields?: IListViewFieldConfig[];
  formFields?: IListViewFieldConfig[];
}

export default class TkmsWebpartWebPart extends BaseClientSideWebPart<ITkmsWebpartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ITkmsWebpartProps> = React.createElement(
      TkmsWebpart,
      {
        description: this.properties.description,
        newButtonText: this.properties.newButtonText || 'New',
        listId: this.properties.listId,
        listViewFields: this.properties.listViewFields || [],
        formFields: this.properties.formFields || [],
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        userDisplayName: this.context.pageContext.user.displayName
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('newButtonText', {
                  label: 'New Button Text'
                }),
                PropertyFieldListPicker('listId', {
                  label: strings.ListPickerFieldLabel,
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  context: this.context as any,
                  selectedList: this.properties.listId,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'listPickerField'
                }),
                PropertyFieldCollectionData('listViewFields', {
                  key: 'listViewFields',
                  label: strings.ListViewFieldsLabel,
                  panelHeader: strings.ListViewFieldsPanelHeader,
                  manageBtnLabel: strings.ListViewFieldsManageBtnLabel,
                  value: this.properties.listViewFields || [],
                  fields: [
                    {
                      id: 'FieldInternalName',
                      title: strings.ListViewFieldInternalName,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'Title'
                    },
                    {
                      id: 'FieldDisplayName',
                      title: strings.ListViewFieldDisplayName,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'Title'
                    },
                    {
                      id: 'FieldName',
                      title: strings.ListViewFieldName,
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'Title'
                    },
                    {
                      id: 'FieldType',
                      title: strings.ListViewFieldType,
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: [
                        { key: 'Text', text: 'Text' },
                        { key: 'Number', text: 'Number' },
                        { key: 'DateTime', text: 'DateTime' },
                        { key: 'Boolean', text: 'Boolean' },
                        { key: 'Choice', text: 'Choice' },
                        { key: 'Lookup', text: 'Lookup' },
                        { key: 'Person', text: 'Person' }
                      ]
                    },
                    {
                      id: 'Width',
                      title: strings.ListViewFieldWidth,
                      type: CustomCollectionFieldType.number,
                      required: true,
                      placeholder: 'Width'
                    },
                  ]
                }),
                PropertyFieldCollectionData('formFields', {
                  key: 'formFields',
                  label: 'Form Fields',
                  panelHeader: 'Configure Form Fields',
                  manageBtnLabel: 'Manage Form Fields',
                  value: this.properties.formFields || [],
                  fields: [
                    {
                      id: 'FieldInternalName',
                      title: 'Field Internal Name',
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'Internal Name'
                    },
                    {
                      id: 'FieldDisplayName',
                      title: 'Field Display Name',
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'Display Name'
                    },
                    {
                      id: 'FieldType',
                      title: 'Field Type',
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: [
                        { key: 'Text', text: 'Text' },
                        { key: 'Number', text: 'Number' },
                        { key: 'DateTime', text: 'DateTime' },
                        { key: 'Boolean', text: 'Boolean' },
                        { key: 'Choice', text: 'Choice' },
                        { key: 'Lookup', text: 'Lookup' },
                        { key: 'Person', text: 'Person' },
                        { key: 'CrossSiteLookup', text: 'CrossSiteLookup' }
                      ]
                    },
                    {
                      id: 'FieldConfig',
                      title: 'Field Configuration',
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'Configuration'
                    }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
