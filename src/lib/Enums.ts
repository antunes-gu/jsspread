const ValueInputOption = Object.freeze({
  raw: 'RAW',
  user_entered: 'USER_ENTERED'
});

const InsertDataOption = Object.freeze({
  overwrite: 'OVERWRITE',
  insert_rows: 'INSERT_ROWS'
});

const ValueRenderOption = Object.freeze({
  formatted: 'FORMATTED_VALUE',
  unformatted: 'UNFORMATTED_VALUE',
  formula: 'FORMULA'
});

const ExportFormat = Object.freeze({
  pdf: 'pdf',
  excel: 'xlsx',
  csv: 'csv',
  open_office_sheet: 'ods',
  tsv: 'tsv',
  zipped_html: 'zip'
});