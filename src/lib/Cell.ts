/**
 * Inicializa o objeto da célula.
 * @class
 */
class Cell {
  declare row: number;
  declare col: number;
  declare value: any;

  /**
   * Representa uma célula individual na planilha.
   * 
   * @param {number} row - O número da linha (começa em 1).
   * @param {number} col - O número da coluna (começa em 1).
   * @param {any} value - O valor contido na célula.
   */
  constructor(row: number, col: number, value: any) {
    this.row = row;
    this.col = col;
    this.value = value;
  }

  /**
   * Retorna a notação A1 da célula (ex: linha 1, coluna 1 vira "A1").
   * @returns {string}
   */
  get address(): string {
    let tempCol = this.col;
    let letter = '';
    
    while (tempCol > 0) {
      let remainder = (tempCol - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      tempCol = Math.floor((tempCol - 1) / 26);
    }
    
    return `${letter}${this.row}`;
  }

  /**
   * Retorna o valor da célula convertido para número, ou null se não for possível.
   * @returns {number|null}
   */
  get numeric_value(): number | null {
    if (this.value === null || this.value === undefined || this.value === '') {
      return null;
    }
    const parsed = Number(this.value);
    return isNaN(parsed) ? null : parsed;
  }

  /**
   * Cria uma nova instância de Cell a partir de uma notação A1 e um valor.
   * Método estático (chamado direto na classe: Cell.fromAddress('B5', 10)).
   * 
   * @param {string} label - A notação A1 (ex: "B5").
   * @param {any} value - O valor da célula.
   * @returns {Cell} Uma nova instância de Cell.
   */
  static fromAddress(label: string, value: any): Cell {
    // Ex: "AB123" -> match[1] = "AB", match[2] = "123"
    const match = label.toUpperCase().match(/^([A-Z]+)([0-9]+)$/);

    if (!match) {
      throw new Error(`ValueError: Notação A1 inválida '${label}'.`);
    }

    const colStr = match[1];
    const row = parseInt(match[2], 10);

    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
      col = col * 26 + (colStr.charCodeAt(i) - 64);
    }

    return new Cell(row, col, value);
  }

  /**
   * Retorna uma representação em string da célula para debug.
   * @returns {string}
   */
  toString(): string {
    return `Cell(${this.address}, '${this.value}')`;
  }

  /**
   * Converte a célula em um objeto JSON puro.
   * @returns {Record<string, any>}
   */
  toJSON(): Record<string, any> {
    return {
      row: this.row,
      col: this.col,
      address: this.address,
      value: this.value
    };
  }
}
