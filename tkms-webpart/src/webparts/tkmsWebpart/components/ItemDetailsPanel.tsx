import * as React from 'react';
import {
  Panel,
  PanelType,
  Stack,
  TextField,
  Dropdown,
  DatePicker,
  Toggle,
  Text,
  Label,
  PrimaryButton,
  DefaultButton,
  Dialog,
  DialogType,
  DialogFooter,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import type { IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/controls/peoplepicker/IPeoplePickerContext';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import type { IPersonaProps } from '@fluentui/react/lib/Persona';
import type { IListViewFieldConfig } from '../TkmsWebpartWebPart';
import styles from './TkmsWebpart.module.scss';

interface IItemDetailsPanelProps {
  isOpen: boolean;
  isNewItemMode?: boolean;
  onClose: () => void;
  onItemDeleted?: () => void;
  onItemSaved?: () => void;
  selectedItem: { [key: string]: string | number | boolean | undefined } | undefined;
  fields: IListViewFieldConfig[];
  spHttpClient?: SPHttpClient;
  siteUrl?: string;
  listId?: string;
}

interface IItemDetailsPanelState {
  formValues: { [key: string]: string | number | boolean | undefined };
  isEditMode: boolean;
  originalValues: { [key: string]: string | number | boolean | undefined };
  isSaving: boolean;
  messageState: {
    type: MessageBarType | undefined;
    text: string;
  };
  isDeleteDialogOpen: boolean;
  choiceOptions: { [fieldName: string]: Array<{ key: string; text: string }> };
  multiChoiceFields: { [fieldName: string]: boolean };
}

export default class ItemDetailsPanel extends React.Component<IItemDetailsPanelProps, IItemDetailsPanelState> {
  public constructor(props: IItemDetailsPanelProps) {
    super(props);

    this.state = {
      formValues: {},
      isEditMode: false,
      originalValues: {},
      isSaving: false,
      messageState: {
        type: undefined,
        text: ''
      },
      isDeleteDialogOpen: false,
      choiceOptions: {},
      multiChoiceFields: {}
    };
  }

  public componentDidMount(): void {
    this._loadChoiceOptions().catch(() => undefined);
  }

  public componentDidUpdate(prevProps: IItemDetailsPanelProps): void {
    if (this.props.isOpen && (!prevProps.isOpen || prevProps.isNewItemMode !== this.props.isNewItemMode) && this.props.isNewItemMode) {
      this._initializeNewItemForm();
      return;
    }

    if (prevProps.selectedItem !== this.props.selectedItem && this.props.selectedItem) {
      this.setState({
        formValues: { ...this.props.selectedItem },
        originalValues: { ...this.props.selectedItem },
        isEditMode: false,
        messageState: { type: undefined, text: '' }
      }, () => {
        this._loadChoiceOptions().catch(() => undefined);
      });
    }
  }

  private _initializeNewItemForm = (): void => {
    const initialValues: { [key: string]: string | number | boolean | undefined } = {};

    for (const field of this.props.fields || []) {
      const fieldName = field.FieldInternalName;
      if (!fieldName || fieldName === 'ID') {
        continue;
      }

      if (field.FieldType?.toLowerCase() === 'boolean') {
        initialValues[fieldName] = false;
      } else {
        initialValues[fieldName] = '';
      }
    }

    this.setState({
      formValues: initialValues,
      originalValues: { ...initialValues },
      isEditMode: true,
      messageState: { type: undefined, text: '' }
    }, () => {
      this._loadChoiceOptions().catch(() => undefined);
    });
  };

  private _handleFieldChange = (fieldName: string, value: string | number | boolean | undefined): void => {
    this.setState(prevState => ({
      formValues: {
        ...prevState.formValues,
        [fieldName]: value
      }
    }));
  };

  private _enterEditMode = (): void => {
    this.setState({ isEditMode: true, messageState: { type: undefined, text: '' } });
  };

  private _cancelEdit = (): void => {
    if (this.props.isNewItemMode) {
      this.props.onClose();
      return;
    }

    this.setState({
      isEditMode: false,
      formValues: { ...this.state.originalValues },
      messageState: { type: undefined, text: '' }
    });
  };

  private _hasChanged = (): boolean => {
    const { formValues, originalValues } = this.state;
    return JSON.stringify(formValues) !== JSON.stringify(originalValues);
  };

  private _normalizeChoiceValue = (value: string | number | boolean | undefined): string => {
    if (value === undefined || value === null) {
      return '';
    }

    if (typeof value === 'object') {
      const objectValue = value as { Key?: string | number; key?: string | number; Text?: string; text?: string; Value?: string; value?: string };
      return String(
        objectValue.Key ?? objectValue.key ?? objectValue.Text ?? objectValue.text ?? objectValue.Value ?? objectValue.value ?? ''
      );
    }

    return String(value);
  };

  private _formatDisplayDate = (value: string | number | boolean | undefined): string => {
    if (value === undefined || value === null || value === '') {
      return '';
    }

    const dateValue = new Date(String(value));
    if (Number.isNaN(dateValue.getTime())) {
      return String(value);
    }

    const month = (dateValue.getMonth() + 1).toString();
    const day = dateValue.getDate().toString();
    const year = dateValue.getFullYear().toString();

    return `${month.length === 1 ? `0${month}` : month}/${day.length === 1 ? `0${day}` : day}/${year}`;
  };

  private _getPanelHeaderText = (): string => {
    const { fields, selectedItem, isNewItemMode } = this.props;
    const { formValues } = this.state;

    if (isNewItemMode) {
      return 'New Item';
    }

    if (!fields || fields.length === 0) {
      return 'Item Details';
    }

    const firstField = fields.find(field => field.FieldInternalName && field.FieldInternalName !== 'ID');
    if (!firstField || !firstField.FieldInternalName) {
      return 'Item Details';
    }

    const label = firstField.FieldDisplayName || firstField.FieldName || firstField.FieldInternalName;
    const sourceValue = formValues[firstField.FieldInternalName] ?? selectedItem?.[firstField.FieldInternalName];

    if (sourceValue === undefined || sourceValue === null || sourceValue === '') {
      return `${label} : -`;
    }

    if (firstField.FieldType?.toLowerCase() === 'datetime') {
      return `${label} : ${this._formatDisplayDate(sourceValue)}`;
    }

    if (typeof sourceValue === 'boolean') {
      return `${label} : ${sourceValue ? 'Yes' : 'No'}`;
    }

    return `${label} : ${this._normalizeChoiceValue(sourceValue)}`;
  };

  private _loadChoiceOptions = async (): Promise<void> => {
    const { spHttpClient, siteUrl, listId, fields } = this.props;

    if (!spHttpClient || !siteUrl || !listId || !fields) {
      return;
    }

    const choiceOptions: { [fieldName: string]: Array<{ key: string; text: string }> } = {};
    const multiChoiceFields: { [fieldName: string]: boolean } = {};

    try {
      for (const field of fields) {
        if (field.FieldType?.toLowerCase() === 'choice') {
          const fieldInternalName = field.FieldInternalName;
          const fieldsUrl = `${siteUrl}/_api/web/lists(guid'${listId}')/fields?$filter=InternalName eq '${encodeURIComponent(fieldInternalName)}'`;

          const response: SPHttpClientResponse = await spHttpClient.get(
            fieldsUrl,
            SPHttpClient.configurations.v1
          );

          if (response.ok) {
            const data = await response.json() as {
              value?: Array<{
                Choices?: string[];
                TypeAsString?: string;
                AllowMultipleValues?: boolean;
              }>;
            };
            if (data.value && data.value.length > 0) {
              const fieldDef = data.value[0];
              const choices: string[] = fieldDef.Choices || [];
              choiceOptions[fieldInternalName] = choices.map(choice => ({
                key: choice,
                text: choice
              }));

              const typeAsString = (fieldDef.TypeAsString || '').toLowerCase();
              multiChoiceFields[fieldInternalName] = typeAsString === 'multichoice' || fieldDef.AllowMultipleValues === true;
            }
          }
        } else if (field.FieldType?.toLowerCase() === 'crosssitelookup') {
          const fieldInternalName = field.FieldInternalName;
          const configuration = (field.FieldConfig || '').trim();
          const separatorIndex = configuration.indexOf(',');
          if (!fieldInternalName || separatorIndex < 0) {
            console.warn(`CrossSiteLookup field '${fieldInternalName}' requires FieldConfig in the format WebUrl,ListTitle.`);
            continue;
          }

          const webUrl = configuration.slice(0, separatorIndex).trim().replace(/\/$/, '');
          const listTitle = configuration.slice(separatorIndex + 1).trim();
          if (!webUrl || !listTitle) {
            console.warn(`CrossSiteLookup field '${fieldInternalName}' has an incomplete FieldConfig.`);
            continue;
          }

          try {
            let nextUrl: string | undefined = `${webUrl}/_api/web/lists/getbytitle('${listTitle.replace(/'/g, "''")}')/items?$select=Id,Title&$top=5000`;
            const options: Array<{ key: string; text: string }> = [];
            const seenTitles = new Set<string>();

            while (nextUrl) {
              const response: SPHttpClientResponse = await spHttpClient.get(nextUrl, SPHttpClient.configurations.v1);
              if (!response.ok) {
                throw new Error(`Failed to load options from '${listTitle}': ${response.status} ${response.statusText}`);
              }

              const data = await response.json() as {
                value?: Array<{ Title?: string }>;
                d?: { results?: Array<{ Title?: string }>; __next?: string };
                '@odata.nextLink'?: string;
                'odata.nextLink'?: string;
              };
              const items = data.value || data.d?.results || [];
              items.forEach(item => {
                const title = item.Title?.trim();
                if (title && !seenTitles.has(title)) {
                  seenTitles.add(title);
                  options.push({ key: title, text: title });
                }
              });

              const pageLink = data['@odata.nextLink'] || data['odata.nextLink'] || data.d?.__next;
              nextUrl = pageLink ? new URL(pageLink, webUrl).toString() : undefined;
            }

            choiceOptions[fieldInternalName] = options;
          } catch (error) {
            console.error(`Error loading CrossSiteLookup options for '${fieldInternalName}':`, error);
          }
        }
      }

      this.setState({ choiceOptions, multiChoiceFields });
    } catch (error) {
      console.error('Error loading choice options:', error);
    }
  };

  private _resolvePersonFieldId = async (fieldName: string, fieldValue: string | number | boolean | undefined): Promise<number | null> => {
    const { spHttpClient, siteUrl } = this.props;
    if (!spHttpClient || !siteUrl) {
      return null;
    }

    const rawValue = fieldValue === undefined || fieldValue === null ? '' : String(fieldValue).trim();
    if (!rawValue) {
      return null;
    }

    const numericValue = Number(rawValue);
    if (!Number.isNaN(numericValue) && numericValue > 0) {
      return numericValue;
    }

    try {
      const loginName = rawValue;
      if (rawValue.indexOf('@') > -1) {
        const emailUrl = `${siteUrl}/_api/web/siteusers/getByEmail('${encodeURIComponent(rawValue)}')`;
        const emailResponse = await spHttpClient.get(emailUrl, SPHttpClient.configurations.v1);
        if (emailResponse.ok) {
          const emailData = await emailResponse.json();
          if (emailData && emailData.Id) {
            return Number(emailData.Id);
          }
        }
      }

      const ensureUserUrl = `${siteUrl}/_api/web/ensureUser`;
      const ensureResponse = await spHttpClient.post(
        ensureUserUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({ logonName: loginName })
        }
      );

      if (!ensureResponse.ok) {
        return null;
      }

      const ensureData = await ensureResponse.json();
      return ensureData && ensureData.Id ? Number(ensureData.Id) : null;
    } catch (error) {
      console.error(`Failed to resolve person field '${fieldName}'`, error);
      return null;
    }
  };

  private _saveItem = async (): Promise<void> => {
    const { spHttpClient, siteUrl, listId, onItemSaved } = this.props;
    const { formValues, originalValues } = this.state;

    if (!spHttpClient || !siteUrl || !listId) {
      this.setState({
        messageState: {
          type: MessageBarType.error,
          text: 'Missing required configuration'
        }
      });
      return;
    }

    this.setState({ isSaving: true });

    try {
      const itemId = formValues.ID as number;
      if (!itemId) {
        throw new Error('Item ID not found');
      }

      const updateUrl = `${siteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})`;
      const updatePayload: { [key: string]: unknown } = {};

      for (const field of this.props.fields || []) {
        const fieldName = field.FieldInternalName;
        if (!fieldName || fieldName === 'ID') {
          continue;
        }

        const fieldValue = formValues[fieldName];
        const originalValue = originalValues[fieldName];
        const isChanged = JSON.stringify(fieldValue ?? null) !== JSON.stringify(originalValue ?? null);

        if (!isChanged) {
          continue;
        }

        if (field.FieldType?.toLowerCase() === 'person') {
          const resolvedPersonId = await this._resolvePersonFieldId(fieldName, fieldValue);
          updatePayload[`${fieldName}Id`] = resolvedPersonId ?? null;
          continue;
        }

        if (field.FieldType?.toLowerCase() === 'choice' && this.state.multiChoiceFields[fieldName]) {
          const multiChoiceValues = String(fieldValue ?? '')
            .split(',')
            .map(value => value.trim())
            .filter(Boolean);
          updatePayload[fieldName] = multiChoiceValues;
          continue;
        }

        updatePayload[fieldName] = fieldValue ?? '';
      }

      if (Object.keys(updatePayload).length === 0) {
        this.setState({
          isEditMode: false,
          messageState: {
            type: MessageBarType.info,
            text: 'No changes to save'
          },
          isSaving: false
        });
        return;
      }

      const response: SPHttpClientResponse = await spHttpClient.post(
        updateUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'X-HTTP-Method': 'MERGE',
            'If-Match': '*'
          },
          body: JSON.stringify(updatePayload)
        }
      );

      if (!response.ok) {
        const errorBody = await response.text();
        throw new Error(`Failed to save item: ${response.status} ${response.statusText}${errorBody ? ` - ${errorBody}` : ''}`);
      }

      this.setState({
        isEditMode: false,
        originalValues: { ...formValues },
        messageState: {
          type: MessageBarType.success,
          text: 'Item saved successfully'
        },
        isSaving: false
      });

      if (onItemSaved) {
        onItemSaved();
      }

      setTimeout(() => {
        this.setState({
          messageState: { type: undefined, text: '' }
        });
      }, 3000);
    } catch (error) {
      this.setState({
        messageState: {
          type: MessageBarType.error,
          text: error instanceof Error ? error.message : 'Error saving item'
        },
        isSaving: false
      });
    }
  };

  private _createItem = async (): Promise<void> => {
    const { spHttpClient, siteUrl, listId, onItemSaved } = this.props;
    const { formValues } = this.state;

    if (!spHttpClient || !siteUrl || !listId) {
      this.setState({
        messageState: {
          type: MessageBarType.error,
          text: 'Missing required configuration'
        }
      });
      return;
    }

    this.setState({ isSaving: true });

    try {
      const createUrl = `${siteUrl}/_api/web/lists(guid'${listId}')/items`;
      const createPayload: { [key: string]: unknown } = {};

      for (const field of this.props.fields || []) {
        const fieldName = field.FieldInternalName;
        if (!fieldName || fieldName === 'ID') {
          continue;
        }

        const fieldValue = formValues[fieldName];
        const isEmptyValue = fieldValue === undefined || fieldValue === null || fieldValue === '';

        if (field.FieldType?.toLowerCase() === 'person') {
          if (isEmptyValue) {
            continue;
          }

          const resolvedPersonId = await this._resolvePersonFieldId(fieldName, fieldValue);
          if (resolvedPersonId) {
            createPayload[`${fieldName}Id`] = resolvedPersonId;
          }
          continue;
        }

        if (field.FieldType?.toLowerCase() === 'choice' && this.state.multiChoiceFields[fieldName]) {
          const multiChoiceValues = String(fieldValue ?? '')
            .split(',')
            .map(value => value.trim())
            .filter(Boolean);

          if (multiChoiceValues.length > 0) {
            createPayload[fieldName] = multiChoiceValues;
          }
          continue;
        }

        if (field.FieldType?.toLowerCase() === 'boolean') {
          createPayload[fieldName] = fieldValue === true || fieldValue === 'true' || fieldValue === 'Yes';
          continue;
        }

        if (isEmptyValue) {
          continue;
        }

        createPayload[fieldName] = fieldValue;
      }

      const response: SPHttpClientResponse = await spHttpClient.post(
        createUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(createPayload)
        }
      );

      if (!response.ok) {
        const errorBody = await response.text();
        throw new Error(`Failed to create item: ${response.status} ${response.statusText}${errorBody ? ` - ${errorBody}` : ''}`);
      }

      this.setState({
        isSaving: false,
        isEditMode: false,
        messageState: {
          type: MessageBarType.success,
          text: 'Item created successfully'
        }
      });

      if (onItemSaved) {
        onItemSaved();
      }

      this.props.onClose();
    } catch (error) {
      this.setState({
        messageState: {
          type: MessageBarType.error,
          text: error instanceof Error ? error.message : 'Error creating item'
        },
        isSaving: false
      });
    }
  };

  private _deleteItem = async (): Promise<void> => {
    const { spHttpClient, siteUrl, listId } = this.props;
    const { formValues, isSaving } = this.state;

    if (isSaving) {
      return;
    }

    if (!spHttpClient || !siteUrl || !listId) {
      this.setState({
        messageState: {
          type: MessageBarType.error,
          text: 'Missing required configuration'
        }
      });
      return;
    }

    const itemId = Number(formValues.ID);
    if (!itemId) {
      this.setState({
        messageState: {
          type: MessageBarType.error,
          text: 'Item ID not found'
        }
      });
      return;
    }

    this.setState({ isDeleteDialogOpen: true });
  };

  private _confirmDeleteItem = async (): Promise<void> => {
    const { spHttpClient, siteUrl, listId, onItemDeleted } = this.props;
    const { formValues } = this.state;

    const itemId = Number(formValues.ID);
    if (!itemId || !spHttpClient || !siteUrl || !listId) {
      this.setState({
        isDeleteDialogOpen: false,
        messageState: {
          type: MessageBarType.error,
          text: 'Unable to move item to Recycle Bin'
        }
      });
      return;
    }

    this.setState({ isDeleteDialogOpen: false, isSaving: true, messageState: { type: undefined, text: '' } });

    try {
      const recycleUrl = `${siteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/recycle()`;

      const response: SPHttpClientResponse = await spHttpClient.post(
        recycleUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json'
          }
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to move item to Recycle Bin: ${response.status} ${response.statusText}`);
      }

      this.setState({ isSaving: false, isEditMode: false });
      if (onItemDeleted) {
        onItemDeleted();
      }
      this.props.onClose();
    } catch (error) {
      this.setState({
        messageState: {
          type: MessageBarType.error,
          text: error instanceof Error ? error.message : 'Error moving item to Recycle Bin'
        },
        isSaving: false
      });
    }
  };

  private _dismissDeleteDialog = (): void => {
    this.setState({ isDeleteDialogOpen: false });
  };

  private _renderFormField = (field: IListViewFieldConfig, index: number): React.ReactNode => {
    const { formValues, isEditMode } = this.state;
    const fieldValue = formValues[field.FieldInternalName] ?? '';
    const fieldType = field.FieldType?.toLowerCase() || 'text';
    const isFieldDisabled = field.FieldInternalName === 'ID' || !isEditMode;

    const labelStyles = {
      root: {
        backgroundColor: '#003767',
        color: '#ffffff',
        fontWeight: 600 as const,
        padding: '8px 12px',
        borderRadius: '4px 0 0 4px',
        display: 'flex',
        alignItems: 'center',
        minWidth: '150px'
      }
    };

    const horizontalStackStyles = { root: { marginBottom: '0px', display: 'flex', border: '1px solid #e1e1e1' } };

    switch (fieldType) {
      case 'text':
      case 'string':
        return (
          <Stack key={index} horizontal styles={horizontalStackStyles}>
            <Label styles={labelStyles}>{field.FieldDisplayName || field.FieldName}</Label>
            <TextField
              value={String(fieldValue)}
              onChange={(_, newValue) => !isFieldDisabled && this._handleFieldChange(field.FieldInternalName, newValue)}
              readOnly={isFieldDisabled}
              borderless={true}
              styles={{ root: { flex: 1 } }}
            />
          </Stack>
        );

      case 'number':
        return (
          <Stack key={index} horizontal styles={horizontalStackStyles}>
            <Label styles={labelStyles}>{field.FieldDisplayName || field.FieldName}</Label>
            <TextField
              type="number"
              value={String(fieldValue)}
              onChange={(_, newValue) => !isFieldDisabled && this._handleFieldChange(field.FieldInternalName, newValue)}
              readOnly={isFieldDisabled}
              borderless={true}
              styles={{ root: { flex: 1 } }}
            />
          </Stack>
        );

      case 'datetime':
        const dateValue = fieldValue ? new Date(String(fieldValue)) : undefined;
        const formattedDateValue = this._formatDisplayDate(fieldValue as string | number | boolean | undefined);
        return (
          <Stack key={index} horizontal styles={horizontalStackStyles}>
            <Label styles={labelStyles}>{field.FieldDisplayName || field.FieldName}</Label>
            {isEditMode ? (
              <DatePicker
                value={dateValue}
                formatDate={(date) => {
                  if (!date) {
                    return '';
                  }
                  const month = (date.getMonth() + 1).toString();
                  const day = date.getDate().toString();
                  const year = date.getFullYear().toString();
                  return `${month.length === 1 ? `0${month}` : month}/${day.length === 1 ? `0${day}` : day}/${year}`;
                }}
                onSelectDate={(date) => !isFieldDisabled && this._handleFieldChange(field.FieldInternalName, date?.toISOString() ?? '')}
                disabled={isFieldDisabled}
                styles={{ root: { flex: 1 } }}
              />
            ) : (
              <TextField
                value={formattedDateValue}
                readOnly={true}
                borderless={true}
                styles={{ root: { flex: 1 } }}
              />
            )}
          </Stack>
        );

      case 'boolean':
        return (
          <Stack key={index} horizontal styles={horizontalStackStyles}>
            <Label styles={labelStyles}>{field.FieldDisplayName || field.FieldName}</Label>
            <Toggle
              checked={fieldValue === true || fieldValue === 'true' || fieldValue === 'Yes'}
              disabled={isFieldDisabled}
              onText="Yes"
              offText="No"
              onChange={(_, checked) => !isFieldDisabled && this._handleFieldChange(field.FieldInternalName, checked)}
              styles={{ root: { flex: 1, marginLeft: '12px' } }}
            />
          </Stack>
        );

      case 'choice':
        const { choiceOptions } = this.state;
        const currentValue = this._normalizeChoiceValue(fieldValue as string | number | boolean | undefined);
        const options = choiceOptions[field.FieldInternalName] || [];
        const dropdownOptions = options.some(option => String(option.key) === currentValue)
          ? options
          : currentValue
            ? [...options, { key: currentValue, text: currentValue }]
            : options;
        const selectedKey = dropdownOptions.some(option => String(option.key) === currentValue)
          ? currentValue
          : undefined;

        if (!isEditMode) {
          return (
            <Stack key={index} horizontal styles={horizontalStackStyles}>
              <Label styles={labelStyles}>{field.FieldDisplayName || field.FieldName}</Label>
              <TextField
                value={currentValue || ''}
                readOnly={true}
                borderless={true}
                styles={{ root: { flex: 1 } }}
              />
            </Stack>
          );
        }

        return (
          <Stack key={index} horizontal styles={horizontalStackStyles}>
            <Label styles={labelStyles}>{field.FieldDisplayName || field.FieldName}</Label>
            <Dropdown
              key={`${field.FieldInternalName}-${currentValue}-${dropdownOptions.length}`}
              selectedKey={selectedKey}
              options={dropdownOptions}
              disabled={isFieldDisabled}
              onChange={(_, option) => !isFieldDisabled && option && this._handleFieldChange(field.FieldInternalName, String(option.key))}
              styles={{ root: { flex: 1 } }}
            />
          </Stack>
        );

      case 'crosssitelookup':
        const crossSiteCurrentValue = this._normalizeChoiceValue(fieldValue as string | number | boolean | undefined);
        const crossSiteOptions = this.state.choiceOptions[field.FieldInternalName] || [];
        const crossSiteDropdownOptions = crossSiteOptions.some(option => option.key === crossSiteCurrentValue)
          ? crossSiteOptions
          : crossSiteCurrentValue
            ? [...crossSiteOptions, { key: crossSiteCurrentValue, text: crossSiteCurrentValue }]
            : crossSiteOptions;

        if (!isEditMode) {
          return (
            <Stack key={index} horizontal styles={horizontalStackStyles}>
              <Label styles={labelStyles}>{field.FieldDisplayName || field.FieldName}</Label>
              <TextField
                value={crossSiteCurrentValue}
                readOnly={true}
                borderless={true}
                styles={{ root: { flex: 1 } }}
              />
            </Stack>
          );
        }

        return (
          <Stack key={index} horizontal styles={horizontalStackStyles}>
            <Label styles={labelStyles}>{field.FieldDisplayName || field.FieldName}</Label>
            <Dropdown
              selectedKey={crossSiteCurrentValue || undefined}
              options={crossSiteDropdownOptions}
              disabled={isFieldDisabled}
              onChange={(_, option) => !isFieldDisabled && option && this._handleFieldChange(field.FieldInternalName, String(option.key))}
              styles={{ root: { flex: 1 } }}
            />
          </Stack>
        );

      case 'lookup':
        return (
          <Stack key={index} horizontal styles={horizontalStackStyles}>
            <Label styles={labelStyles}>{field.FieldDisplayName || field.FieldName}</Label>
            <TextField
              value={String(fieldValue)}
              onChange={(_, newValue) => !isFieldDisabled && this._handleFieldChange(field.FieldInternalName, newValue)}
              readOnly={isFieldDisabled}
              borderless={true}
              styles={{ root: { flex: 1 } }}
            />
          </Stack>
        );

      case 'person':
        const personValue = this._normalizeChoiceValue(fieldValue as string | number | boolean | undefined);
          const peoplePickerContext = this.props.siteUrl && this.props.spHttpClient
          ? {
              absoluteUrl: this.props.siteUrl,
            msGraphClientFactory: undefined as unknown as IPeoplePickerContext['msGraphClientFactory'],
              spHttpClient: this.props.spHttpClient
            }
          : undefined;

        if (!isEditMode) {
          return (
            <Stack key={index} horizontal styles={horizontalStackStyles}>
              <Label styles={labelStyles}>{field.FieldDisplayName || field.FieldName}</Label>
              <Text variant="small" styles={{ root: { flex: 1, padding: '8px 12px', border: '1px solid #e1e1e1', minHeight: '32px', display: 'flex', alignItems: 'center' } }}>
                {personValue || ''}
              </Text>
            </Stack>
          );
        }

        return (
          <Stack key={index} horizontal styles={horizontalStackStyles}>
            <Label styles={labelStyles}>{field.FieldDisplayName || field.FieldName}</Label>
            {peoplePickerContext ? (
              <PeoplePicker
                context={peoplePickerContext as unknown as IPeoplePickerContext}
                placeholder="Search people"
                personSelectionLimit={1}
                defaultSelectedUsers={personValue ? [personValue] : []}
                principalTypes={[PrincipalType.User]}
                ensureUser={true}
                onChange={(selectedUsers: IPersonaProps[]) => {
                  const selectedUser = selectedUsers[0] as IPersonaProps & { loginName?: string };
                  const selectedPerson = selectedUsers && selectedUsers.length > 0
                    ? (selectedUser.loginName || selectedUser.secondaryText || selectedUser.text || selectedUser.id || '')
                    : '';
                  this._handleFieldChange(field.FieldInternalName, selectedPerson);
                }}
                disabled={isFieldDisabled}
                styles={{ root: { flex: 1 } }}
              />
            ) : (
              <TextField
                value={personValue}
                readOnly={true}
                borderless={true}
                styles={{ root: { flex: 1 } }}
              />
            )}
          </Stack>
        );

      default:
        return (
          <Stack key={index} horizontal styles={horizontalStackStyles}>
            <Label styles={labelStyles}>{field.FieldDisplayName || field.FieldName}</Label>
            <TextField
              value={String(fieldValue)}
              onChange={(_, newValue) => !isFieldDisabled && this._handleFieldChange(field.FieldInternalName, newValue)}
              readOnly={isFieldDisabled}
              borderless={true}
              styles={{ root: { flex: 1 } }}
            />
          </Stack>
        );
    }
  };

  public render(): React.ReactElement<IItemDetailsPanelProps> {
    const { isOpen, onClose, fields, isNewItemMode } = this.props;
    const { isEditMode, messageState, isSaving, isDeleteDialogOpen } = this.state;
    const hasChanges = this._hasChanged();
    const panelHeaderText = this._getPanelHeaderText();

    const onRenderFooterContent = (): JSX.Element | null => {
      if (isEditMode) {
        return (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              className={styles.panelPrimaryButton}
              text={isNewItemMode ? 'Create' : 'Save'}
              onClick={isNewItemMode ? this._createItem : this._saveItem}
              disabled={!hasChanges || isSaving}
            />
            <DefaultButton
              text="Cancel"
              onClick={this._cancelEdit}
              disabled={isSaving}
            />
          </Stack>
        );
      }
      return null;
    };

    return (
      <Panel
        isOpen={isOpen}
        onDismiss={onClose}
        type={PanelType.medium}
        headerText={panelHeaderText}
        closeButtonAriaLabel="Close"
        onRenderFooterContent={onRenderFooterContent}
        isFooterAtBottom={true}
        styles={{
          navigation: {
            background: 'linear-gradient(90deg, #bb181d 0%, #003767 100%)'
          },
          header: {
            background: 'transparent',
            paddingTop: '16px',
            paddingBottom: '16px'
          },
          headerText: {
            color: '#ffffff'
          },
          commands: {
            selectors: {
              '.ms-Button': {
                color: '#ffffff'
              },
              '.ms-Button-icon': {
                color: '#ffffff'
              },
              '.ms-Button:hover': {
                backgroundColor: 'rgba(255, 255, 255, 0.12)'
              },
              '.ms-Button:hover .ms-Button-icon': {
                color: '#ffffff'
              }
            }
          }
        }}
      >
        <Stack styles={{ root: { marginTop: '5px', paddingTop: '5px', borderTop: '1px solid #edebe9' } }}>
        <Stack styles={{ root: { minHeight: '32px', marginBottom: '16px' } }}>
          {messageState.type ? (
            <MessageBar messageBarType={messageState.type}>
              {messageState.text}
            </MessageBar>
          ) : null}
        </Stack>

        <Stack tokens={{ childrenGap: 12 }}>
          {!isNewItemMode && (
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton
                className={styles.panelPrimaryButton}
                text={isEditMode ? 'Editing' : 'Edit'}
                onClick={this._enterEditMode}
                disabled={isSaving || isEditMode}
                onRenderText={() => (
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                    {isEditMode && <Spinner size={SpinnerSize.xSmall} />}
                    <span>{isEditMode ? 'Editing' : 'Edit'}</span>
                  </Stack>
                )}
              />
              <DefaultButton
                text="Delete"
                onClick={this._deleteItem}
                disabled={isSaving}
              />
            </Stack>
          )}

          <Stack tokens={{ childrenGap: 0 }}>
            {fields && fields.length > 0 ? (
              fields.map((field, index) => this._renderFormField(field, index))
            ) : (
              <Text variant="small">No fields configured.</Text>
            )}
          </Stack>
        </Stack>
        </Stack>

        <Dialog
          hidden={!isDeleteDialogOpen}
          onDismiss={this._dismissDeleteDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Confirm Delete',
            subText: 'Are you sure you want to move this item to the Recycle Bin?'
          }}
        >
          <DialogFooter>
            <PrimaryButton className={styles.panelPrimaryButton} text="Delete" onClick={this._confirmDeleteItem} />
            <DefaultButton text="Cancel" onClick={this._dismissDeleteDialog} />
          </DialogFooter>
        </Dialog>
      </Panel>
    );
  }
}
