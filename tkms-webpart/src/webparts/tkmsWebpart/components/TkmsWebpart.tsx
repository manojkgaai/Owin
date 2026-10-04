import * as React from 'react';
import styles from './TkmsWebpart.module.scss';
import type { ITkmsWebpartProps } from './ITkmsWebpartProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { DetailsList, IColumn, SelectionMode, Stack, Text, Link, PrimaryButton, Spinner, SpinnerSize } from '@fluentui/react';
import ItemDetailsPanel from './ItemDetailsPanel';

interface IListItemState {
  items: Array<{ [key: string]: string | number | boolean | undefined }>;
  columns: IColumn[];
  loading: boolean;
  error: string;
  isPanelOpen: boolean;
  isNewItemMode: boolean;
  selectedItem: { [key: string]: string | number | boolean | undefined } | undefined;
}

export default class TkmsWebpart extends React.Component<ITkmsWebpartProps, IListItemState> {
  public constructor(props: ITkmsWebpartProps) {
    super(props);

    this.state = {
      items: [],
      columns: [],
      loading: false,
      error: '',
      isPanelOpen: false,
      isNewItemMode: false,
      selectedItem: undefined
    };
  }

  public componentDidMount(): void {
    this._loadListData().catch(() => undefined);
  }

  public componentDidUpdate(prevProps: ITkmsWebpartProps): void {
    if (prevProps.listId !== this.props.listId || prevProps.listViewFields !== this.props.listViewFields || prevProps.siteUrl !== this.props.siteUrl) {
      this._loadListData().catch(() => undefined);
    }
  }

  private _openPanel = (item: { [key: string]: string | number | boolean | undefined }): void => {
    this.setState({ isPanelOpen: true, selectedItem: item, isNewItemMode: false });
  };

  private _openNewItemPanel = (): void => {
    this.setState({ isPanelOpen: true, selectedItem: {}, isNewItemMode: true });
  };

  private _closePanel = (): void => {
    this.setState({ isPanelOpen: false, selectedItem: undefined, isNewItemMode: false });
  };

  private _handleItemDeleted = (): void => {
    this.setState({ isPanelOpen: false, selectedItem: undefined, isNewItemMode: false });
    this._loadListData().catch(() => undefined);
  };

  private _handleItemSaved = (): void => {
    this._loadListData(true).catch(() => undefined);
  };

  private _formatCsvValue(value: unknown, fieldType?: string): string {
    if (value === undefined || value === null) {
      return '';
    }

    if (Array.isArray(value)) {
      return value.map(item => this._formatCsvValue(item, fieldType)).filter(Boolean).join(', ');
    }

    if (typeof value === 'object') {
      const item = value as { [key: string]: unknown };
      const displayValue = item.Title ?? item.LookupValue ?? item.Name ?? item.EMail ?? item.Email;
      if (displayValue !== undefined && displayValue !== null) {
        return String(displayValue);
      }

      if (Array.isArray(item.results)) {
        return this._formatCsvValue(item.results, fieldType);
      }

      return this._formatValue(value, fieldType);
    }

    return this._formatValue(value, fieldType);
  }

  private _downloadCsv = async (): Promise<void> => {
    const { listId, formFields, spHttpClient, siteUrl } = this.props;
    const fields = (formFields || []).filter(field => field.FieldInternalName);

    if (!listId || fields.length === 0 || !spHttpClient || !siteUrl) {
      return;
    }

    const expandedFields = fields.filter(field => {
      const type = (field.FieldType || '').toLowerCase();
      return type === 'person' || type === 'lookup';
    });
    const selectedFields = fields.map(field => {
      const type = (field.FieldType || '').toLowerCase();
      return type === 'person' || type === 'lookup'
        ? `${field.FieldInternalName}/Id,${field.FieldInternalName}/Title`
        : field.FieldInternalName;
    }).join(',');
    const expandQuery = expandedFields.length > 0
      ? `&$expand=${encodeURIComponent(expandedFields.map(field => field.FieldInternalName).join(','))}`
      : '';

    let nextUrl: string | undefined = `${siteUrl}/_api/web/lists(guid'${listId}')/items?$select=${encodeURIComponent(selectedFields)}${expandQuery}&$top=5000`;
    const listItems: Array<{ [key: string]: unknown }> = [];

    try {
      while (nextUrl) {
        const response: SPHttpClientResponse = await spHttpClient.get(nextUrl, SPHttpClient.configurations.v1);
        if (!response.ok) {
          throw new Error(`Failed to download list items: ${response.status} ${response.statusText}`);
        }

        const data = await response.json() as {
          value?: Array<{ [key: string]: unknown }>;
          d?: { results?: Array<{ [key: string]: unknown }>; __next?: string };
          '@odata.nextLink'?: string;
          'odata.nextLink'?: string;
        };
        listItems.push(...(data.value || data.d?.results || []));

        const pageLink = data['@odata.nextLink'] || data['odata.nextLink'] || data.d?.__next;
        nextUrl = pageLink ? new URL(pageLink, siteUrl).toString() : undefined;
      }

      const escapeCsvCell = (value: string): string => `"${value.replace(/"/g, '""')}"`;
      const header = fields.map(field => escapeCsvCell(
        field.FieldDisplayName || field.FieldName || field.FieldInternalName
      ));
      const rows = listItems.map(item => fields.map(field => escapeCsvCell(
        this._formatCsvValue(item[field.FieldInternalName], field.FieldType)
      )));
      const csvContent = `\uFEFF${[header, ...rows].map(row => row.join(',')).join('\r\n')}`;
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const downloadUrl = URL.createObjectURL(blob);
      const downloadLink = document.createElement('a');
      downloadLink.href = downloadUrl;
      downloadLink.download = `sharepoint-list-${new Date().toISOString().slice(0, 10)}.csv`;
      document.body.appendChild(downloadLink);
      downloadLink.click();
      document.body.removeChild(downloadLink);
      URL.revokeObjectURL(downloadUrl);
    } catch (error) {
      this.setState({ error: error instanceof Error ? error.message : 'Unable to download CSV.' });
    }
  };

  private _formatValue(value: unknown, fieldType?: string): string {
    if (value === undefined || value === null) {
      return '';
    }

    if (Array.isArray(value)) {
      return value
        .map(item => this._formatValue(item, fieldType))
        .filter(Boolean)
        .join(', ');
    }

    if (typeof value === 'object') {
      const obj = value as { [key: string]: unknown };
      const title = typeof obj.Title === 'string' ? obj.Title : typeof obj.title === 'string' ? obj.title : '';
      const name = typeof obj.Name === 'string' ? obj.Name : typeof obj.name === 'string' ? obj.name : '';
      const email = typeof obj.EMail === 'string' ? obj.EMail : typeof obj.Email === 'string' ? obj.Email : '';

      const firstNamePart = (input: string): string => {
        if (!input) {
          return '';
        }

        const trimmed = input.trim();
        if (!trimmed) {
          return '';
        }

        const firstToken = trimmed.split(/\s+/)[0];
        return firstToken || trimmed;
      };

      if (title) {
        return firstNamePart(title);
      }

      if (name) {
        return firstNamePart(name);
      }

      if (email) {
        return email;
      }

      const nested = Object.keys(obj)
        .map(key => this._formatValue((obj as { [key: string]: unknown })[key], fieldType))
        .filter(value => value && value.length > 0)
        .join(', ');

      return nested;
    }

    if (fieldType && fieldType.toLowerCase() === 'datetime') {
      const date = new Date(value as string | number | Date);
      if (Number.isNaN(date.getTime())) {
        return String(value);
      }

      const month = date.getMonth() + 1;
      const day = date.getDate();
      const year = date.getFullYear();
      const monthText = month < 10 ? `0${month}` : String(month);
      const dayText = day < 10 ? `0${day}` : String(day);
      return `${monthText}/${dayText}/${year}`;
    }

    if (typeof value === 'boolean') {
      return value ? 'Yes' : 'No';
    }

    return String(value);
  }

  private _normalizeFieldValue(value: unknown, fieldType?: string): string | number | boolean | undefined {
    if (value === undefined || value === null) {
      return undefined;
    }

    if (Array.isArray(value)) {
      const joined = value
        .map(item => this._formatValue(item, fieldType))
        .filter(Boolean)
        .join(', ');

      return joined;
    }

    if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
      return value;
    }

    return this._formatValue(value, fieldType);
  }

  private async _loadListData(silentRefresh?: boolean): Promise<void> {
    const {
      listId,
      listViewFields,
      formFields,
      spHttpClient,
      siteUrl
    } = this.props;

    if (!listId || !listViewFields || listViewFields.length === 0 || !spHttpClient || !siteUrl) {
      this.setState({ items: [], columns: [], loading: false, error: '' });
      return;
    }

    const allFields = [...(listViewFields || []), ...(formFields || [])];
    const uniqueFields = allFields.filter((field, index, arr) =>
      field.FieldInternalName && arr.findIndex(item => item.FieldInternalName === field.FieldInternalName) === index
    );

    const personFields = uniqueFields.filter(field => field.FieldType && field.FieldType.toLowerCase() === 'person');

    const selectedFields = uniqueFields
      .map(field => {
        const fieldName = field.FieldInternalName;
        if (!fieldName) {
          return '';
        }

        if (personFields.some(personField => personField.FieldInternalName === fieldName)) {
          return `${fieldName}/Id,${fieldName}/Title`;
        }

        return fieldName;
      })
      .filter(Boolean)
      .join(',');

    // Always include ID field for save operations
    const fieldsWithId = selectedFields ? `ID,${selectedFields}` : 'ID';

    if (!fieldsWithId) {
      this.setState({ items: [], columns: [], loading: false, error: '' });
      return;
    }

    const expandedFields = personFields
      .map(field => field.FieldInternalName)
      .filter(Boolean)
      .join(',');

    const expandQuery = expandedFields ? `&$expand=${encodeURIComponent(expandedFields)}` : '';
    let nextUrl: string | undefined = `${siteUrl}/_api/web/lists(guid'${listId}')/items?$select=${encodeURIComponent(fieldsWithId)}${expandQuery}&$orderby=ID%20asc&$top=5000`;

    if (!silentRefresh) {
      this.setState({ loading: true, error: '' });
    } else {
      this.setState({ error: '' });
    }

    try {
      const rawItems: Array<{ [key: string]: unknown }> = [];
      while (nextUrl) {
        const response: SPHttpClientResponse = await spHttpClient.get(nextUrl, SPHttpClient.configurations.v1);

        if (!response.ok) {
          throw new Error(`Failed to load list items: ${response.status} ${response.statusText}`);
        }

        const data = await response.json() as {
          value?: Array<{ [key: string]: unknown }>;
          d?: { results?: Array<{ [key: string]: unknown }>; __next?: string };
          '@odata.nextLink'?: string;
          'odata.nextLink'?: string;
        };
        rawItems.push(...(data.value || data.d?.results || []));

        const pageLink = data['@odata.nextLink'] || data['odata.nextLink'] || data.d?.__next;
        nextUrl = pageLink ? new URL(pageLink, siteUrl).toString() : undefined;
      }

      const columns: IColumn[] = listViewFields.map((field, index) => {
        const width = Number(field.Width ?? 140);
        const columnWidth = Number.isFinite(width) && width > 0 ? width : 140;

        return {
          key: field.FieldInternalName || `field-${index}`,
          name: field.FieldDisplayName || field.FieldName || field.FieldInternalName || `Field ${index + 1}`,
          fieldName: field.FieldInternalName,
          minWidth: columnWidth,
          maxWidth: columnWidth,
          width: columnWidth,
          isResizable: true,
          onRender: (item: { [key: string]: string | number | boolean | undefined }): React.ReactNode => {
            const value = this._formatValue(item[field.FieldInternalName], field.FieldType);
            if (index === 0) {
              return (
                <Link onClick={() => this._openPanel(item)}>
                  {value}
                </Link>
              );
            }
            return <span>{value}</span>;
          }
        };
      });

      const normalizedItems = rawItems.map((item: { [key: string]: unknown }) => {
        const result: { [key: string]: string | number | boolean | undefined } = {};
        
        // Include ID and all configured fields
        if (item.ID) {
          result.ID = item.ID as number;
        }
        
        uniqueFields.forEach(field => {
          const fieldName = field.FieldInternalName;
          const value = item[fieldName];
          result[fieldName] = this._normalizeFieldValue(value, field.FieldType);
        });
        return result;
      });

      this.setState({ items: normalizedItems, columns, loading: false, error: '' });
    } catch (error) {
      this.setState({
        items: [],
        columns: [],
        loading: false,
        error: error instanceof Error ? error.message : 'Unable to load list data.'
      });
    }
  }

  public render(): React.ReactElement<ITkmsWebpartProps> {
    const {
      listId,
      listViewFields,
      newButtonText,
    } = this.props;

    const { items, columns, loading, error, isPanelOpen, isNewItemMode, selectedItem } = this.state;
    const configuredNewButtonText = (newButtonText || '').trim();
    const normalizedNewButtonText = configuredNewButtonText.toLowerCase() === 'new' ? '' : configuredNewButtonText;
    const newButtonLabel = normalizedNewButtonText ? `New ${normalizedNewButtonText}` : 'New';

    return (
      <section className={`${styles.tkmsWebpart}`}>
        <div className={styles.welcome}>
          <Stack tokens={{ childrenGap: 12 }}>
            {error ? (
              <Text variant="small" styles={{ root: { color: '#a4262c' } }}>{error}</Text>
            ) : null}

            {loading ? (
              <Spinner label="Loading list data..." size={SpinnerSize.small} />
            ) : listId && listViewFields && listViewFields.length > 0 ? (
              <Stack tokens={{ childrenGap: 10 }}>
                <Stack horizontal tokens={{ childrenGap: 8 }}>
                  <PrimaryButton
                    text={newButtonLabel}
                    onClick={this._openNewItemPanel}
                    iconProps={{ iconName: 'Add' }}
                    styles={{
                      root: {
                        minWidth: '96px',
                        background: 'linear-gradient(90deg, #bb181d 0%, #003767 100%)',
                        border: 'none'
                      },
                      rootHovered: {
                        background: 'linear-gradient(90deg, #a31519 0%, #002e57 100%)',
                        border: 'none'
                      },
                      rootPressed: {
                        background: 'linear-gradient(90deg, #8f1216 0%, #002643 100%)',
                        border: 'none'
                      },
                      label: { color: '#ffffff' },
                      icon: { color: '#ffffff' }
                    }}
                  />
                  <PrimaryButton
                    text="Download CSV"
                    onClick={this._downloadCsv}
                    iconProps={{ iconName: 'Download' }}
                    disabled={!this.props.formFields || this.props.formFields.length === 0}
                    styles={{
                      root: {
                        minWidth: '96px',
                        background: 'linear-gradient(90deg, #bb181d 0%, #003767 100%)',
                        border: 'none'
                      },
                      rootHovered: {
                        background: 'linear-gradient(90deg, #a31519 0%, #002e57 100%)',
                        border: 'none'
                      },
                      rootPressed: {
                        background: 'linear-gradient(90deg, #8f1216 0%, #002643 100%)',
                        border: 'none'
                      },
                      label: { color: '#ffffff' },
                      icon: { color: '#ffffff' }
                    }}
                  />
                </Stack>
                <DetailsList
                  className={styles.detailsList}
                  items={items}
                  columns={columns}
                  compact={true}
                  selectionMode={SelectionMode.none}
                  isHeaderVisible={true}
                  styles={{ root: { overflowX: 'auto' } }}
                />
              </Stack>
            ) : (
              <Text variant="small">No list data to display.</Text>
            )}
          </Stack>
        </div>

        <ItemDetailsPanel
          isOpen={isPanelOpen}
          onClose={this._closePanel}
          onItemDeleted={this._handleItemDeleted}
          onItemSaved={this._handleItemSaved}
          isNewItemMode={isNewItemMode}
          selectedItem={selectedItem}
          fields={this.props.formFields || []}
          spHttpClient={this.props.spHttpClient}
          siteUrl={this.props.siteUrl}
          listId={this.props.listId}
        />
      </section>
    );
  }
}