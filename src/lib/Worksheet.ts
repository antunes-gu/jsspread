/**
 * Inicializa o objeto da aba.
 * @class
 */
class Worksheet {
  spreadsheet: Spreadsheet;
  sheet: GoogleAppsScript.Spreadsheet.Sheet;

  /**
   * Inicializa o objeto da Aba.
   * 
   * OBSERVAÇÃO: Assim como a Spreadsheet, esta classe é um wrapper. 
   * A aba nativa do Google Apps Script fica armazenada no atributo `this.sheet`.
   * Você pode acessar os métodos nativos usando `ws.sheet.getRange(...)`, etc.
   * 
   * @param {Spreadsheet} spreadsheet - A instância da classe Spreadsheet principal.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} gasSheet - O objeto nativo da aba do Google Apps Script.
   */
  constructor(spreadsheet: Spreadsheet, gasSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    this.spreadsheet = spreadsheet;
    this.sheet = gasSheet;
  }

  // ==========================================
  // Propriedades da Worksheet
  // ==========================================
  
  /**
   * Retorna a URL direta da aba.
   * @returns {string}
   */
  get url(): string {
    const spreadsheetUrl = this.sheet.getParent().getUrl();
    const sheetId = this.sheet.getSheetId();
    return `${spreadsheetUrl}#gid=${sheetId}`;
  }


  /**
   * Retorna o ID único (GID) da aba.
   * @returns {string}
   */
  get id(): string {
    return this.sheet.getSheetId().toString();
  }

  /**
   * Retorna o nome da aba (como se fosse um atributo).
   */
  get title(): string {
    return this.sheet.getName();
  }

  /**
   * Retorna o índice numérico da aba na planilha (começa em 0 ou 1 dependendo da API, 
   * no GAS getIndex() começa em 1).
   * @returns {number}
   */
  get index(): number {
    return this.sheet.getIndex();
  }

  /**
   * Retorna o número total de linhas alocadas na aba (Não conta as linhas em branco).
   * @returns {number}
   */
  get row_count(): number {
    return this.sheet.getLastRow();
  }

  /**
   * Retorna o número total de colunas alocadas na aba (Não conta as colunas em branco).
   * @returns {number}
   */
  get column_count(): number {
    return this.sheet.getLastColumn();
  }
  
  // ==========================================
  // Métodos Públicos
  // ==========================================
  
  /**
   * Exporta exclusivamente esta aba no formato especificado.
   * 
   * @param {string} [format=ExportFormat.pdf] - O formato de exportação (use ExportFormat).
   * @returns {GoogleAppsScript.Base.Blob} O arquivo exportado.
   */
  export(format: string = ExportFormat.pdf): GoogleAppsScript.Base.Blob {
    return this.spreadsheet.export(format, this);
  }

  /**
   * Retorna todos os valores de uma linha específica.
   * @param {number} index - O número da linha (começa em 1).
   * @returns {any[]}
   */
  rowValues(index: number): any[] {
    const lastCol = this.sheet.getLastColumn();
    if (lastCol === 0) return [];

    const data = this.sheet.getRange(index, 1, 1, lastCol).getValues();
    return data[0];
  }

  /**
   * Retorna todos os valores de uma coluna específica.
   * @param {number} index - O número da coluna (começa em 1).
   * @returns {any[]}
   */
  colValues(index: number): any[] {
    const lastRow = this.sheet.getLastRow();
    if (lastRow === 0) return [];

    const data = this.sheet.getRange(1, index, lastRow, 1).getValues();
    return data.map(row => row[0]);
  }

  /**
   * Retorna os valores de um intervalo específico da planilha.
   * 
   * @param {string} range - Intervalo em notação A1 (ex: "A1:C5").
   * @returns {any[][]} Uma matriz bidimensional com os valores do intervalo.
   */
  getWithRange(range: string): any[][] {
    return this.sheet.getRange(range).getValues();
  }

  /**
   * Retorna todos os valores da planilha como uma lista de listas.
   * @returns {any[][]}
   */
  getAllValues(): any[][] {
    const data = this.sheet.getDataRange().getValues();
    if (data.length === 0 || (data.length === 1 && data[0][0] === "")) return [];

    return data;
  }

  /**
   * Retorna todos os valores da planilha como uma lista de objetos mapeados pelo cabeçalho.
   * 
   * @param {number} head - A linha que contém o cabeçalho (começa em 1).
   * @param {boolean} empty2zero - Converte valores em branco para 0.
   * @param {any} default_blank - Valor padrão para células em branco (caso empty2zero seja false).
   * 
   * @returns {Record<string, any>[]}
   */
  getAllRecords(head: number = 1, empty2zero: boolean = false, default_blank: any = ""): Record<string, any>[] {
    const data = this.sheet.getDataRange().getValues();
    if (data.length < head) return [];

    const headers = data[head - 1];
    const rows = data.slice(head);

    return rows.map(row => {
      let record: Record<string, any> = {};
      headers.forEach((header: string, index: number) => {
        if (header !== "") {
          let value = row[index];

          if (value === "") {
            if (empty2zero) {
              value = 0;
            } else {
              value = default_blank;
            }
          }

          record[header] = value;
        }
      });
      return record;
    });
  }

  /**
   * Insere uma linha de dados em uma posição específica, empurrando as demais para baixo.
   * 
   * @param {number} index - A posição onde a linha será inserida (começa em 1).
   * @param {any[]} values - Uma lista simples com os dados da linha.
   * @param {Record<string, any>} [options] - Configurações opcionais.
   * @returns {Worksheet} Retorna a própria instância da classe.
   */
  insertRow(index: number, values: any[], options: Record<string, any> = {}): Worksheet {
    this._isSimpleArray(values, "TypeError: insertRow aceita apenas uma lista simples. Para múltiplas linhas, use insertRows()");
    const valueInputOption = this._verifyInputOption(options.valueInputOption);

    this.sheet.insertRowBefore(index);

    const targetRange = this.sheet.getRange(index, 1, 1, values.length);
    this._applyValues(targetRange, [values], valueInputOption);

    return this;
  }

  /**
   * Insere múltiplas linhas de dados em uma posição específica.
   * 
   * @param {number} index - A posição inicial onde as linhas serão inseridas (começa em 1).
   * @param {any[][]} rows - Matriz (lista de listas) com os dados.
   * @param {Record<string, any>} [options] - Configurações opcionais.
   * @returns {Worksheet} Retorna a própria instância da classe.
   */
  insertRows(index: number, rows: any[][], options: Record<string, any> = {}): Worksheet {
    this._isValidMatrix(rows, "TypeError: insertRows exige uma lista de listas (matriz). Para uma única linha, use insertRow()");
    const valueInputOption = this._verifyInputOption(options.valueInputOption);

    const rowCount = rows.length;
    const maxColumns = Math.max(...rows.map(r => Array.isArray(r) ? r.length : 0));

    this.sheet.insertRowsBefore(index, rowCount);

    const targetRange = this.sheet.getRange(index, 1, rowCount, maxColumns);
    const normalizedMatrix = this._normalizeMatrix(rows, maxColumns);

    this._applyValues(targetRange, normalizedMatrix, valueInputOption);

    return this;
  }

  /**
   * Adiciona uma única linha de dados ao final da planilha.
   * 
   * @param {any[]} values - Uma lista simples com os dados da linha (ex: ['João', 25]).
   * @param {Record<string, any>} [options] - Configurações opcionais.
   * @returns {Worksheet} Retorna a própria instância da classe.
   */
  appendRow(values: any[], options: Record<string, any> = {}): Worksheet {
    this._isSimpleArray(values, "TypeError: appendRow aceita apenas uma lista simples. Para múltiplas linhas, use appendRows().");

    const valueInputOption = this._verifyInputOption(options.valueInputOption);
    const insertDataOption = this._verifyInsertOption(options.insertDataOption);

    const lastRow = this.sheet.getLastRow();

    if (insertDataOption === InsertDataOption.insert_rows && lastRow > 0) {
      this.sheet.insertRowAfter(lastRow);
    }

    const targetRow = lastRow + 1;
    const targetRange = this.sheet.getRange(targetRow, 1, 1, values.length);

    this._applyValues(targetRange, [values], valueInputOption);

    return this;
  }

  /**
   * Adiciona múltiplas linhas de dados de uma só vez ao final da planilha.
   * 
   * @param {any[][]} rows - Uma lista de listas contendo as linhas (ex: [['Maria', 30], ['Pedro', 22]]).
   * @param {Record<string, any>} [options] - Configurações opcionais.
   * @returns {Worksheet} Retorna a própria instância da classe.
   */
  appendRows(rows: any[][], options: Record<string, any> = {}): Worksheet {
    this._isValidMatrix(rows, "TypeError: appendRows exige uma lista de listas (matriz). Para uma única linha, use appendRow().");

    const valueInputOption = this._verifyInputOption(options.valueInputOption);
    const insertDataOption = this._verifyInsertOption(options.insertDataOption);

    const rowCount = rows.length;
    const maxColumns = Math.max(...rows.map(r => Array.isArray(r) ? r.length : 0));
    const lastRow = this.sheet.getLastRow();

    if (insertDataOption === InsertDataOption.insert_rows && lastRow > 0) {
      this.sheet.insertRowsAfter(lastRow, rowCount);
    }

    const startRow = lastRow + 1;
    const targetRange = this.sheet.getRange(startRow, 1, rowCount, maxColumns);

    const normalizedMatrix = this._normalizeMatrix(rows, maxColumns);

    this._applyValues(targetRange, normalizedMatrix, valueInputOption);

    return this;
  }

  /**
   * Atualiza um intervalo de células com novos valores.
   * 
   * @param {string} range - Intervalo em notação A1 (ex: "A1" ou "A1:B2")
   * @param {any[][]} values - Matriz de valores (lista de listas).
   * @param {Record<string, any>} [options] - Configurações opcionais.
   * @returns {Worksheet} Retorna a própria instância da classe.
   */
  update(range: string, values: any[][], options: Record<string, any> = {}): Worksheet {
    this._isValidMatrix(values, "TypeError: O método update exige uma matriz de valores (lista de listas).");

    const valueInputOption = this._verifyInputOption(options.valueInputOption);

    const numRows = values.length;
    const maxCols = Math.max(...values.map(r => Array.isArray(r) ? r.length : 0));

    const initialRange = this.sheet.getRange(range);
    const startRow = initialRange.getRow();
    const startCol = initialRange.getColumn();

    const targetRange = this.sheet.getRange(startRow, startCol, numRows, maxCols);

    const normalizedMatrix = this._normalizeMatrix(values, maxCols);
    
    this._applyValues(targetRange, normalizedMatrix, valueInputOption);

    return this;
  }

  /**
   * Atualiza o valor de uma célula específica passando linha e coluna.
   * 
   * @param {number} row - O número da linha (começa em 1).
   * @param {number} col - O número da coluna (começa em 1).
   * @param {any} value - O valor a ser inserido na célula.
   * @param {Record<string, any>} [options] - Configurações opcionais.
   * @returns {Worksheet} Retorna a própria instância da classe.
   */
  updateCell(row: number, col: number, value: any, options: Record<string, any> = {}): Worksheet {
    const valueInputOption = this._verifyInputOption(options.valueInputOption);
    const cellRange = this.sheet.getRange(row, col);

    this._applyValues(cellRange, value, valueInputOption);

    return this;
  }

  /**
   * Atualiza o valor de uma célula específica usando notação A1.
   * 
   * @param {string} range - Intervalo em notação A1 (ex: "A1" ou "A1:B2")
   * @param {any} value - O valor a ser inserido na célula.
   * @param {Record<string, any>} [options] - Configurações opcionais.
   * @returns {Worksheet} Retorna a própria instância da classe.
   */
  updateACell(range: string, value: any, options: Record<string, any> = {}): Worksheet {
    const valueInputOption = this._verifyInputOption(options.valueInputOption);
    const cellRange = this.sheet.getRange(range);
    
    this._applyValues(cellRange, value, valueInputOption);

    return this;
  }

  /**
   * Retorna um objeto Cell correspondente a uma célula específica usando notação A1.
   * 
   * @param {string} label - A notação A1 da célula (ex: "B1").
   * @param {Record<string, any>} [options] - Configurações opcionais.
   * @returns {any} Instância da classe Cell contendo a linha, coluna e o valor solicitado.
   */
  acell(label: string, options: Record<string, any> = {}): any { // Assumindo que o tipo Cell seja implementado depois
    const renderOption = this._verifyRenderOption(options.valueRenderOption);

    const range = this.sheet.getRange(label);
    const row = range.getRow();
    const col = range.getColumn();

    let cellValue;

    if (renderOption === ValueRenderOption.formula) {
      cellValue = range.getFormula() || range.getValue();
    } else {
      cellValue = renderOption === ValueRenderOption.formatted 
                  ? range.getDisplayValue() 
                  : range.getValue();
    }

    return new Cell(row, col, cellValue);
  }

  /**
   * Retorna um objeto Cell correspondente a uma célula específica usando linha e coluna.
   * 
   * @param {number} row - O número da linha (começa em 1).
   * @param {number} col - O número da coluna (começa em 1).
   * @param {Record<string, any>} [options] - Configurações opcionais.
   * @returns {any} Instância da classe Cell contendo a linha, coluna e o valor solicitado.
   */
  cell(row: number, col: number, options: Record<string, any> = {}): any { // Assumindo Cell
    const renderOption = this._verifyRenderOption(options.valueRenderOption);
    
    const range = this.sheet.getRange(row, col);
    
    let cellValue;

    // A correção de bug no nome da constante foi necessária aqui também, pois no código original estava em MAIUSCÚLO (ValueRenderOption.FORMULA) 
    if (renderOption === ValueRenderOption.formula) {
      cellValue = range.getFormula() || range.getValue();
    } else {
      cellValue = renderOption === ValueRenderOption.formatted 
                  ? range.getDisplayValue() 
                  : range.getValue();
    }
    
    // @ts-ignore
    return new Cell(row, col, cellValue);
  }

  /**
   * Encontra a primeira célula que contém o valor ou que corresponda ao Regex.
   * 
   * @param {string|number|RegExp} query - O valor ou expressão regular a ser buscada.
   * @returns {any|null} Retorna uma instância da classe Cell, ou null se não encontrar.
   */
  find(query: string | number | RegExp): any | null {
    if (query === undefined || query === null) return null;

    const isRegex = query instanceof RegExp;

    // Se for Regex, pega apenas o padrão em texto (a propriedade .source)
    const searchString = isRegex ? query.source : query.toString();

    const textFinder = this.sheet.createTextFinder(searchString);

    if (isRegex) {
      textFinder.useRegularExpression(true);
      // @ts-ignore
      textFinder.matchCase(!query.ignoreCase);
    }

    const result = textFinder.findNext();

    if (!result) {
      return null;
    }
    
    return new Cell(result.getRow(), result.getColumn(), result.getValue());
  }

  /**
   * Encontra todas as células que contêm o valor ou que correspondam ao Regex.
   * 
   * @param {string|number|RegExp} query - O valor ou expressão regular a ser buscada.
   * @returns {any[]} Uma lista de instâncias da classe Cell.
   */
  findAll(query: string | number | RegExp): any[] {
    if (query === undefined || query === null) return [];

    const isRegex = query instanceof RegExp;
    const searchString = isRegex ? query.source : query.toString();

    const textFinder = this.sheet.createTextFinder(searchString);

    if (isRegex) {
      textFinder.useRegularExpression(true);
      // @ts-ignore
      textFinder.matchCase(!query.ignoreCase);
    }

    const results = textFinder.findAll();
    
    if (results.length === 0) {
      return [];
    }
    
    return results.map(range => new Cell(range.getRow(), range.getColumn(), range.getValue()));
  }

  /**
   * Exclui uma ou mais linhas da planilha fisicamente.
   * 
   * @param {number} index - O número da linha a ser excluída (começa em 1).
   * @param {number} [numRows=1] - A quantidade de linhas a excluir a partir do index.
   * @returns {Worksheet} Retorna a própria instância da classe.
   */
  deleteRows(index: number, numRows: number = 1): Worksheet {
    this.sheet.deleteRows(index, numRows);
    return this;
  }

  /**
   * Exclui uma ou mais colunas da planilha fisicamente.
   * 
   * @param {number} index - O número da coluna a ser excluída (começa em 1).
   * @param {number} [numCols=1] - A quantidade de colunas a excluir a partir do index.
   * @returns {Worksheet} Retorna a própria instância da classe.
   */
  deleteColumns(index: number, numCols: number = 1): Worksheet {
    this.sheet.deleteColumns(index, numCols);
    return this;
  }

  /**
   * Limpa várias áreas específicas da planilha de uma só vez.
   * 
   * @param {string[]} ranges - Lista de strings em notação A1 (ex: ["A1:B2", "D5:E10"])
   * @param {boolean} keepFormat - Se true, apaga apenas os valores. Se false, apaga valores e formatação.
   * @returns {Worksheet} Retorna a própria instância da classe.
   */
  batchClear(ranges: string[], keepFormat: boolean = true): Worksheet {
    if (!Array.isArray(ranges) || ranges.length === 0) {
      return this; 
    }

    const rangeList = this.sheet.getRangeList(ranges);

    if (keepFormat) {
      rangeList.clearContent(); 
    } else {
      rangeList.clear(); 
    }

    return this;
  }

  /**
   * Limpa a aba inteira.
   * 
   * @param {boolean} keepFormat - Se true, apaga apenas os valores. Se false, apaga valores e formatação.
   * @returns {Worksheet} Retorna a própria instância da classe.
   */
  clear(keepFormat: boolean = true): Worksheet {
    if (keepFormat) {
      this.sheet.clearContents();
    } else {
      this.sheet.clear();
    }
    
    return this;
  }

  /**
   * Redimensiona a planilha para ter exatamente o número de linhas e colunas especificado.
   * 
   * @param {number} rows - Número total de linhas desejado.
   * @param {number} cols - Número total de colunas desejado.
   * @returns {Worksheet} Retorna a própria instância da classe.
   */
  resize(rows: number, cols: number): Worksheet {
    const currentRows = this.sheet.getMaxRows();
    const currentCols = this.sheet.getMaxColumns();

    if (rows < currentRows) {
      this.sheet.deleteRows(rows + 1, currentRows - rows);
    } else if (rows > currentRows) {
      this.sheet.insertRowsAfter(currentRows, rows - currentRows);
    }

    if (cols < currentCols) {
      this.sheet.deleteColumns(cols + 1, currentCols - cols);
    } else if (cols > currentCols) {
      this.sheet.insertColumnsAfter(currentCols, cols - currentCols);
    }

    return this;
  }

  // ==========================================
  // Métodos Internos (Auxiliares)
  // ==========================================
  
  /**
   * Verifica se o valor fornecido é uma matriz unidimensional (lista simples).
   * 
   * @private
   */
  _isSimpleArray(array: any, errorMsg: string | null = null): boolean {
    if (!Array.isArray(array) || array.length === 0) {
      if (errorMsg) throw new Error(errorMsg);
      return false;
    }
    
    if (Array.isArray(array[0])) {
      if (errorMsg) throw new Error(errorMsg);
      return false;
    }

    return true;
  }

  /**
   * Verifica se o valor fornecido é uma matriz bidimensional (lista de listas).
   * 
   * @private
   */
  _isValidMatrix(matrix: any, errorMsg: string | null = null): boolean {
    if (!Array.isArray(matrix) || matrix.length === 0) {
      if (errorMsg) throw new Error(errorMsg);
      return false;
    }
    
    if (!Array.isArray(matrix[0])) {
      if (errorMsg) throw new Error(errorMsg);
      return false;
    }

    return true;
  }

  /**
   * Padroniza uma matriz bidimensional, garantindo que todas as linhas tenham o mesmo comprimento.
   * 
   * @private
   */
  _normalizeMatrix(matrix: any[][], maxColumns: number): any[][] {
    return matrix.map(row => {
      const paddedRow = Array.isArray(row) ? [...row] : [row];
      while (paddedRow.length < maxColumns) {
        paddedRow.push("");
      }
      return paddedRow;
    });
  }

  /**
   * Aplica dados a um intervalo, lidando com a formatação RAW ou USER_ENTERED.
   * 
   * @private
   */
  _applyValues(range: GoogleAppsScript.Spreadsheet.Range, data: any | any[][], valueInputOption: string): void {
    const isMatrix = Array.isArray(data);

    if (valueInputOption === ValueInputOption.raw) {
      range.setNumberFormat('@');

      if (isMatrix) {
        const stringMatrix = data.map((row: any[]) => 
          row.map(val => (val !== null && val !== undefined) ? val.toString() : "")
        );
        range.setValues(stringMatrix);
      } else {
        const stringValue = (data !== null && data !== undefined) ? data.toString() : "";
        range.setValue(stringValue);
      }
    } else {
      if (isMatrix) {
        range.setValues(data);
      } else {
        range.setValue(data);
      }
    }
  }

  /**
   * Valida e sanitiza a opção de formatação de entrada de valores.
   * 
   * @private
   */
  _verifyInputOption(valueInputOption?: string): string {
    if (!valueInputOption) {
      return ValueInputOption.user_entered;
    }

    const option = valueInputOption.toUpperCase();

    if (!(Object.values(ValueInputOption) as string[]).includes(option)) {
      throw new Error(`ValueError: Opção recebida '${valueInputOption}' é inválida. Use o objeto ValueInputOption.`);
    }

    return option;
  }

  /**
   * Valida e sanitiza a opção de comportamento de inserção física de dados.
   * 
   * @private
   */
  _verifyInsertOption(insertDataOption?: string): string {
    if (!insertDataOption) {
      return InsertDataOption.insert_rows;
    }

    const option = insertDataOption.toUpperCase();

    if (!(Object.values(InsertDataOption) as string[]).includes(option)) {
      throw new Error(`ValueError: Opção recebida '${insertDataOption}' é inválida. Use o objeto InsertDataOption.`);
    }

    return option;
  }

  /**
   * Valida e sanitiza a opção de comportamento de renderização de dados.
   * 
   * @private
   */
  _verifyRenderOption(valueRenderOption?: string): string {
    if (!valueRenderOption) {
      return ValueRenderOption.formatted;
    }

    const option = valueRenderOption.toUpperCase();

    if (!(Object.values(ValueRenderOption) as string[]).includes(option)) {
      throw new Error(`ValueError: Opção recebida '${valueRenderOption}' é inválida. Use o objeto ValueRenderOption.`);
    }

    return option;
  }
}
