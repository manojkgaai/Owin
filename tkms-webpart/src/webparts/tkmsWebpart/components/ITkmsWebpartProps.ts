import { SPHttpClient } from '@microsoft/sp-http';

export interface IListViewFieldConfig {
  FieldInternalName: string;
  FieldDisplayName: string;
  FieldName: string;
  FieldType: string;
  FieldConfig?: string;
  Width?: number | string;
}

export interface ITkmsWebpartProps {
  description: string;
  newButtonText?: string;
  listId?: string;
  listViewFields?: IListViewFieldConfig[];
  formFields?: IListViewFieldConfig[];
  spHttpClient?: SPHttpClient;
  siteUrl?: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
}
