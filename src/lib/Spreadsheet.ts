/**
 * Inicializa o objeto da Planilha.
 * @class
 */
class Spreadsheet {
  id: string | null;
  file: GoogleAppsScript.Spreadsheet.Spreadsheet;

  /**
   * Inicializa o objeto da Planilha.
   * 
   * OBSERVAÇÃO: Esta classe atua como um "wrapper" (camada de abstração) que oferece 
   * métodos customizados inspirados na biblioteca gspread do Python. Como o objeto 
   * nativo do Google Apps Script é armazenado no atributo `this.file`, você ainda pode 
   * utilizar qualquer método original da engine do Google acessando-o diretamente 
   * (exemplo: `sh.file.addEditor('email@teste.com')`).
   * 
   * @param {string | null} [id] - O ID da planilha. Se não fornecido, usa a planilha ativa.
   */
  constructor(id: string | null = null) {
    this.id = id;
    this.file = id 
      ? SpreadsheetApp.openById(id)
      : SpreadsheetApp.getActiveSpreadsheet();
  }

  // ==========================================
  // Propriedades da Spreadsheet
  // ==========================================
  
  /**
   * Retorna a URL direta da planilha.
   * @returns {string}
   */
  get url(): string {
    return this.file.getUrl();
  }

  /**
   * Retorna o título (nome) do arquivo da planilha.
   * @returns {string}
   */
  get title(): string {
    return this.file.getName();
  }

  /**
   * Retorna a 1ª worksheet do arquivo da planilha.
   * @returns {Worksheet}
   */
  get sheet1(): Worksheet {
    const sheets = this.file.getSheets();
    if (sheets.length === 0) throw new Error("WorksheetNotFound: A planilha não possui abas.");
    return new Worksheet(this, sheets[0]);
  }

  // ==========================================
  // Métodos Públicos
  // ==========================================
  
  worksheet(title: string): Worksheet {
    const sheet = this.file.getSheetByName(title);
    if (!sheet) {
      throw new Error(`WorksheetNotFound: Aba '${title}' não encontrada.`);
    }
    return new Worksheet(this, sheet);
  }

  /**
   * Retorna uma lista com todas as abas da planilha.
   * 
   * @param {boolean} [exclude_hidden=false] - Se true, ignora as abas ocultas.
   * @returns {Worksheet[]} Lista de instâncias da classe Worksheet.
   */
  worksheets(exclude_hidden: boolean = false): Worksheet[] {
    let sheets = this.file.getSheets();
    
    if (exclude_hidden) {
      sheets = sheets.filter(sheet => !sheet.isSheetHidden());
    }

    return sheets.map(sheet => new Worksheet(this, sheet));
  }

  getWorksheet(index: number): Worksheet {
    const sheets = this.file.getSheets();
    
    if (index < 0 || index >= sheets.length) {
      throw new Error(`WorksheetNotFound: Não existe aba no índice ${index}. A planilha tem ${sheets.length} aba(s).`);
    }
    
    return new Worksheet(this, sheets[index]);
  }

  getWorksheetById(id: number): Worksheet {
    // Procura a aba pelo ID (getSheetById não existe nativamente no objeto Spreadsheet)
    const sheet = this.file.getSheets().find(s => s.getSheetId() === id);

    if (!sheet) {
      throw new Error(`WorksheetNotFound: Aba com Id '${id}' não encontrada.`);
    }

    return new Worksheet(this, sheet);
  }

  addWorksheet(title: string, rows: number = 1000, cols: number = 26, index: number | null = null): Worksheet {
    let sheet: GoogleAppsScript.Spreadsheet.Sheet;

    try {
      if (index !== null) {
        sheet = this.file.insertSheet(title, index);
      } else {
        sheet = this.file.insertSheet(title);
      }
    } catch (e) {
      throw new Error(`Erro ao criar aba: Já existe uma aba com o nome '${title}'.`);
    }

    // --- Ajusta a quantidade de linhas ---
    const currentRows = sheet.getMaxRows();
    if (rows && rows < currentRows) {
      sheet.deleteRows(rows + 1, currentRows - rows);
    } else if (rows && rows > currentRows) {
      sheet.insertRowsAfter(currentRows, rows - currentRows);
    }

    // --- Ajusta a quantidade de colunas ---
    const currentCols = sheet.getMaxColumns();
    if (cols && cols < currentCols) {
      sheet.deleteColumns(cols + 1, currentCols - cols);
    } else if (cols && cols > currentCols) {
      sheet.insertColumnsAfter(currentCols, cols - currentCols);
    }

    return new Worksheet(this, sheet);
  }

  delWorksheet(worksheet: Worksheet | string): void {
    let sheet: GoogleAppsScript.Spreadsheet.Sheet | null;
    
    if (typeof worksheet === 'string') {
      sheet = this.file.getSheetByName(worksheet);
      if (!sheet) {
        throw new Error(`WorksheetNotFound: Aba '${worksheet}' não encontrada para exclusão.`);
      }
    } else {
      // Assumindo que worksheet é uma instância da classe Worksheet
      sheet = worksheet.sheet;
    }

    if (sheet) {
      this.file.deleteSheet(sheet);
    }
  }

  /**
   * Exporta a planilha (ou uma aba específica) no formato especificado.
   * 
   * @param {string} [format=ExportFormat.pdf] - O formato de exportação (use ExportFormat).
   * @param {Worksheet | null} [worksheet=null] - (Opcional) A aba específica que deseja exportar.
   * @returns {GoogleAppsScript.Base.Blob} O arquivo exportado.
   */
  export(format: string = ExportFormat.pdf, worksheet: Worksheet | null = null): GoogleAppsScript.Base.Blob {
    if (!(Object.values(ExportFormat) as string[]).includes(format)) {
      throw new Error(`ValueError: Formato '${format}' inválido.`);
    }
    
    // CORREÇÃO AQUI: Usando this.file.getId() ao invés do parâmetro solto "id" que dava erro quando inicializado sem ID
    let url = `https://docs.google.com/spreadsheets/d/${this.file.getId()}/export?format=${format}`;
    
    if (worksheet) {
      // Assumindo que Worksheet tem um atributo .id (que retorna o sheetId) ou podemos acessar worksheet.sheet.getSheetId()
      const sheetId = typeof worksheet.id !== 'undefined' ? worksheet.id : worksheet.sheet.getSheetId();
      url += `&gid=${sheetId}`;
    }

    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      throw new Error(`Erro ao exportar planilha: ${response.getContentText()}`);
    }

    const blob = response.getBlob();
    
    const fileName = worksheet 
      ? `${this.file.getName()} - ${worksheet.title || worksheet.sheet.getName()}` 
      : this.file.getName();
      
    blob.setName(`${fileName}.${format}`);
    
    return blob;
  }
}
