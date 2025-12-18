/**
 * ╔════════════════════════════════════════════════════════════════════════════╗
 * ║  CONVERSOR KLINI SAÚDE v2.0 - DEZEMBRO 2025                               ║
 * ║  Cores das tabelas CORRIGIDAS conforme padrão visual Klini                ║
 * ╚════════════════════════════════════════════════════════════════════════════╝
 * 
 * PALETA DE CORES KLINI (CORRETAS):
 * - Cabeçalho tabelas: #1D7874 (verde teal escuro)
 * - Linha alternada clara: #E8F5F3 (verde menta claro)
 * - Linha alternada branca: #FFFFFF (branco)
 * - Texto cabeçalho: #FFFFFF (branco)
 * - Texto dados: #1D3D3A (verde escuro) ou #333333 (cinza escuro)
 */

import PptxGenJS from "pptxgenjs";
import * as XLSX from "xlsx";

// ============================================================================
// PALETA DE CORES KLINI - CORRIGIDA
// ============================================================================
const KLINI_COLORS = {
  // Cores principais
  TEAL_DARK: "1D7874",      // Verde teal escuro (cabeçalhos)
  TEAL_PRIMARY: "199A8E",   // Verde teal primário
  ORANGE: "F4793B",         // Laranja
  YELLOW: "FFB800",         // Amarelo
  
  // Cores das tabelas
  TABLE_HEADER: "1D7874",           // Cabeçalho: verde teal escuro
  TABLE_ROW_LIGHT: "E8F5F3",        // Linha clara: verde menta muito claro
  TABLE_ROW_WHITE: "FFFFFF",        // Linha branca
  TABLE_FIRST_COLUMN: "E8F5F3",     // Primeira coluna (faixa etária)
  
  // Cores de texto
  TEXT_WHITE: "FFFFFF",
  TEXT_DARK: "1D3D3A",              // Verde escuro para texto
  TEXT_GRAY: "333333",
  
  // Alternativas para variação sutil
  ROW_ALT_1: "FFFFFF",              // Branco
  ROW_ALT_2: "F5FAF9",              // Verde muito sutil (quase branco)
  ROW_ALT_3: "E8F5F3",              // Verde menta claro
};

// ============================================================================
// INTERFACES
// ============================================================================
interface TableRow {
  text: string;
  options?: PptxGenJS.TableCellProps;
}

interface ExcelData {
  comCopart: {
    header: string[];
    rows: string[][];
  };
  semCopart: {
    header: string[];
    rows: string[][];
  };
  planosSemCopart: {
    header: string[];
    rows: string[][];
  };
}

// ============================================================================
// FUNÇÕES AUXILIARES
// ============================================================================

/**
 * Retorna a cor de fundo para uma linha da tabela (padrão zebrado)
 * @param rowIndex Índice da linha (0 = primeira linha de dados)
 * @returns Cor hex sem #
 */
function getRowColor(rowIndex: number): string {
  // Alterna entre branco e verde menta claro
  return rowIndex % 2 === 0 ? KLINI_COLORS.ROW_ALT_1 : KLINI_COLORS.TABLE_ROW_LIGHT;
}

/**
 * Cria uma célula de cabeçalho formatada
 */
function createHeaderCell(text: string): PptxGenJS.TableCell {
  return {
    text: text,
    options: {
      fill: { color: KLINI_COLORS.TABLE_HEADER },
      color: KLINI_COLORS.TEXT_WHITE,
      bold: true,
      align: "center",
      valign: "middle",
      fontSize: 9,
      fontFace: "Arial",
    },
  };
}

/**
 * Cria uma célula de dados formatada
 */
function createDataCell(
  text: string,
  rowIndex: number,
  isFirstColumn: boolean = false
): PptxGenJS.TableCell {
  const bgColor = isFirstColumn 
    ? KLINI_COLORS.TABLE_FIRST_COLUMN 
    : getRowColor(rowIndex);
  
  return {
    text: text,
    options: {
      fill: { color: bgColor },
      color: KLINI_COLORS.TEXT_DARK,
      bold: isFirstColumn,
      align: "center",
      valign: "middle",
      fontSize: 8,
      fontFace: "Arial",
    },
  };
}

/**
 * Cria a tabela de Planos (COM ou SEM coparticipação)
 * Formato: PLANO | REGISTRO ANS | VALOR PER CAPITA | FATURA ESTIMADA
 */
function createPlanosTable(
  header: string[],
  rows: string[][]
): PptxGenJS.TableRow[] {
  const tableData: PptxGenJS.TableRow[] = [];

  // Cabeçalho
  const headerRow: PptxGenJS.TableCell[] = header.map((h) => createHeaderCell(h));
  tableData.push(headerRow);

  // Linhas de dados
  rows.forEach((row, rowIndex) => {
    const dataRow: PptxGenJS.TableCell[] = row.map((cell, colIndex) => {
      // Primeira coluna (PLANO) tem fundo verde claro
      return createDataCell(cell, rowIndex, colIndex === 0);
    });
    tableData.push(dataRow);
  });

  return tableData;
}

/**
 * Cria a tabela de Faixas Etárias (valores por faixa)
 * Formato: FAIXA | PLANO1 | PLANO2 | PLANO3 | ...
 */
function createFaixasTable(
  header: string[],
  rows: string[][]
): PptxGenJS.TableRow[] {
  const tableData: PptxGenJS.TableRow[] = [];

  // Cabeçalho
  const headerRow: PptxGenJS.TableCell[] = header.map((h) => createHeaderCell(h));
  tableData.push(headerRow);

  // Linhas de dados (faixas etárias)
  rows.forEach((row, rowIndex) => {
    const dataRow: PptxGenJS.TableCell[] = row.map((cell, colIndex) => {
      // Primeira coluna (FAIXA ETÁRIA) sempre tem fundo verde claro
      const isFirstCol = colIndex === 0;
      return createDataCell(cell, rowIndex, isFirstCol);
    });
    tableData.push(dataRow);
  });

  return tableData;
}

// ============================================================================
// FUNÇÃO PRINCIPAL DE EXPORTAÇÃO
// ============================================================================

export async function exportToPptx(
  excelData: ExcelData,
  outputPath: string = "proposta_klini.pptx"
): Promise<void> {
  const pptx = new PptxGenJS();
  
  // Configuração da apresentação
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "Klini Saúde";
  pptx.company = "Klini Saúde - ANS 42.202-9";
  pptx.title = "Proposta Comercial PME";

  // -------------------------------------------------------------------------
  // SLIDE 1: CAPA
  // -------------------------------------------------------------------------
  const slideCapa = pptx.addSlide();
  slideCapa.background = { color: KLINI_COLORS.TEAL_PRIMARY };
  
  slideCapa.addText("PROPOSTA COMERCIAL PME", {
    x: 0.5,
    y: 2,
    w: 9,
    h: 1,
    fontSize: 36,
    bold: true,
    color: KLINI_COLORS.TEXT_WHITE,
    fontFace: "Arial",
  });
  
  slideCapa.addText("DEZEMBRO 2025", {
    x: 0.5,
    y: 3.2,
    w: 9,
    h: 0.5,
    fontSize: 18,
    color: KLINI_COLORS.TEXT_WHITE,
    fontFace: "Arial",
  });
  
  slideCapa.addText("ANS - Nº 42.202-9", {
    x: 8,
    y: 0.3,
    w: 2,
    h: 0.3,
    fontSize: 10,
    color: KLINI_COLORS.TEXT_WHITE,
    align: "right",
  });

  // -------------------------------------------------------------------------
  // SLIDE 2: TABELA PLANOS SEM COPARTICIPAÇÃO
  // -------------------------------------------------------------------------
  if (excelData.planosSemCopart.rows.length > 0) {
    const slidePlanos = pptx.addSlide();
    slidePlanos.background = { color: "FFFFFF" };
    
    // Título
    slidePlanos.addText("Planos sem Coparticipação", {
      x: 0.5,
      y: 0.3,
      w: 9,
      h: 0.6,
      fontSize: 24,
      bold: true,
      color: KLINI_COLORS.TEAL_PRIMARY,
      fontFace: "Arial",
    });
    
    // ANS no canto
    slidePlanos.addText("ANS - Nº 42.202-9", {
      x: 8,
      y: 0.3,
      w: 2,
      h: 0.3,
      fontSize: 10,
      color: KLINI_COLORS.TEXT_GRAY,
      align: "right",
    });
    
    // Tabela de planos
    const planosTable = createPlanosTable(
      excelData.planosSemCopart.header,
      excelData.planosSemCopart.rows
    );
    
    slidePlanos.addTable(planosTable, {
      x: 0.5,
      y: 1,
      w: 9,
      colW: [3, 2, 2, 2],
      border: { pt: 0.5, color: "CCCCCC" },
      fontFace: "Arial",
    });
  }

  // -------------------------------------------------------------------------
  // SLIDE 3: TABELA FAIXAS ETÁRIAS - SEM COPARTICIPAÇÃO
  // -------------------------------------------------------------------------
  if (excelData.semCopart.rows.length > 0) {
    const slideFaixas = pptx.addSlide();
    slideFaixas.background = { color: "FFFFFF" };
    
    // Título
    slideFaixas.addText("Valores por Faixa Etária - SEM Coparticipação", {
      x: 0.5,
      y: 0.2,
      w: 9,
      h: 0.5,
      fontSize: 20,
      bold: true,
      color: KLINI_COLORS.TEAL_PRIMARY,
      fontFace: "Arial",
    });
    
    // Tabela de faixas
    const faixasTable = createFaixasTable(
      excelData.semCopart.header,
      excelData.semCopart.rows
    );
    
    // Calcular larguras das colunas dinamicamente
    const numCols = excelData.semCopart.header.length;
    const totalWidth = 9.5;
    const firstColWidth = 0.8;
    const remainingWidth = (totalWidth - firstColWidth) / (numCols - 1);
    const colWidths = [firstColWidth, ...Array(numCols - 1).fill(remainingWidth)];
    
    slideFaixas.addTable(faixasTable, {
      x: 0.25,
      y: 0.8,
      w: totalWidth,
      colW: colWidths,
      border: { pt: 0.5, color: "CCCCCC" },
      fontFace: "Arial",
    });
  }

  // -------------------------------------------------------------------------
  // SLIDE 4: TABELA FAIXAS ETÁRIAS - COM COPARTICIPAÇÃO
  // -------------------------------------------------------------------------
  if (excelData.comCopart.rows.length > 0) {
    const slideFaixasCopart = pptx.addSlide();
    slideFaixasCopart.background = { color: "FFFFFF" };
    
    // Título
    slideFaixasCopart.addText("Valores por Faixa Etária - COM Coparticipação", {
      x: 0.5,
      y: 0.2,
      w: 9,
      h: 0.5,
      fontSize: 20,
      bold: true,
      color: KLINI_COLORS.ORANGE,
      fontFace: "Arial",
    });
    
    // Tabela de faixas
    const faixasCopartTable = createFaixasTable(
      excelData.comCopart.header,
      excelData.comCopart.rows
    );
    
    // Calcular larguras das colunas dinamicamente
    const numCols = excelData.comCopart.header.length;
    const totalWidth = 9.5;
    const firstColWidth = 0.8;
    const remainingWidth = (totalWidth - firstColWidth) / (numCols - 1);
    const colWidths = [firstColWidth, ...Array(numCols - 1).fill(remainingWidth)];
    
    slideFaixasCopart.addTable(faixasCopartTable, {
      x: 0.25,
      y: 0.8,
      w: totalWidth,
      colW: colWidths,
      border: { pt: 0.5, color: "CCCCCC" },
      fontFace: "Arial",
    });
  }

  // Salvar arquivo
  await pptx.writeFile(outputPath);
  console.log(`✅ Apresentação salva: ${outputPath}`);
}

// ============================================================================
// EXPORTAÇÃO DEFAULT
// ============================================================================
export default {
  KLINI_COLORS,
  createHeaderCell,
  createDataCell,
  createPlanosTable,
  createFaixasTable,
  exportToPptx,
};
