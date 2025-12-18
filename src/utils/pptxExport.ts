/**
 * ╔════════════════════════════════════════════════════════════════════════════╗
 * ║  CONVERSOR KLINI SAÚDE v2.0 - DEZEMBRO 2025                               ║
 * ║  Cores das tabelas CORRIGIDAS conforme padrão visual Klini                ║
 * ╚════════════════════════════════════════════════════════════════════════════╝
 */

import PptxGenJS from "pptxgenjs";

// ============================================================================
// PALETA DE CORES KLINI - CORRIGIDA
// ============================================================================
const KLINI_COLORS = {
  // Cores principais
  TEAL_DARK: "1D7874",
  TEAL_PRIMARY: "199A8E",
  ORANGE: "F4793B",
  YELLOW: "FFB800",
  
  // Cores das tabelas
  TABLE_HEADER: "1D7874",
  TABLE_ROW_LIGHT: "E8F5F3",
  TABLE_ROW_WHITE: "FFFFFF",
  TABLE_FIRST_COLUMN: "E8F5F3",
  
  // Cores de texto
  TEXT_WHITE: "FFFFFF",
  TEXT_DARK: "1D3D3A",
  TEXT_GRAY: "333333",
};

// ============================================================================
// INTERFACES
// ============================================================================
interface ParsedData {
  companyName: string;
  concessionaire: string;
  broker: string;
  emissionDate: string;
  validityDate: string;
  demographics: any[];
  plansWithCopay: any[];
  plansWithoutCopay: any[];
  ageBasedPricingCopay: any[];
  ageBasedPricingNoCopay: any[];
  demographicsG?: any[];
  plansWithCopayG?: any[];
  plansWithoutCopayG?: any[];
  ageBasedPricingCopayG?: any[];
  ageBasedPricingNoCopayG?: any[];
}

// ============================================================================
// FUNÇÕES AUXILIARES
// ============================================================================

function getRowColor(rowIndex: number): string {
  return rowIndex % 2 === 0 ? KLINI_COLORS.TABLE_ROW_WHITE : KLINI_COLORS.TABLE_ROW_LIGHT;
}

function formatCurrency(value: number): string {
  return value.toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL",
  });
}

function createHeaderCell(text: string): PptxGenJS.TableCell {
  return {
    text: text,
    options: {
      fill: { color: KLINI_COLORS.TABLE_HEADER },
      color: KLINI_COLORS.TEXT_WHITE,
      bold: true,
      align: "center",
      valign: "middle",
      fontSize: 8,
      fontFace: "Arial",
    },
  };
}

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
      fontSize: 7,
      fontFace: "Arial",
    },
  };
}

// ============================================================================
// FUNÇÃO PRINCIPAL DE EXPORTAÇÃO
// ============================================================================

export async function exportToPPTX(
  data: ParsedData,
  coverImage: string | null,
  includeProductosG: boolean = false
): Promise<void> {
  const pptx = new PptxGenJS();
  
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "Klini Saúde";
  pptx.company = "Klini Saúde - ANS 42.202-9";
  pptx.title = "Proposta Comercial PME";

  // =========================================================================
  // SLIDE 1: CAPA
  // =========================================================================
  const slideCapa = pptx.addSlide();
  
  if (coverImage) {
    slideCapa.addImage({
      data: coverImage,
      x: 0,
      y: 0,
      w: "100%",
      h: "100%",
    });
  } else {
    slideCapa.background = { color: KLINI_COLORS.TEAL_PRIMARY };
    slideCapa.addText("PROPOSTA COMERCIAL PME", {
      x: 0.5, y: 2, w: 9, h: 1,
      fontSize: 36, bold: true, color: KLINI_COLORS.TEXT_WHITE, fontFace: "Arial",
    });
    slideCapa.addText("DEZEMBRO 2025", {
      x: 0.5, y: 3.2, w: 9, h: 0.5,
      fontSize: 18, color: KLINI_COLORS.TEXT_WHITE, fontFace: "Arial",
    });
  }

  // =========================================================================
  // SLIDE 2: PLANOS SEM COPARTICIPAÇÃO
  // =========================================================================
  if (data.plansWithoutCopay && data.plansWithoutCopay.length > 0) {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    
    slide.addText("Planos sem Coparticipação", {
      x: 0.3, y: 0.2, w: 9, h: 0.5,
      fontSize: 22, bold: true, color: KLINI_COLORS.TEAL_PRIMARY, fontFace: "Arial",
    });
    
    slide.addText("ANS - Nº 42.202-9", {
      x: 8.2, y: 0.2, w: 1.5, h: 0.3,
      fontSize: 9, color: KLINI_COLORS.TEXT_GRAY, align: "right",
    });

    const tableData: PptxGenJS.TableRow[] = [];
    
    // Header
    tableData.push([
      createHeaderCell("PLANO"),
      createHeaderCell("REGISTRO ANS"),
      createHeaderCell("VALOR PER CAPITA"),
      createHeaderCell("FATURA ESTIMADA"),
    ]);

    // Data rows
    data.plansWithoutCopay.forEach((plan, idx) => {
      tableData.push([
        createDataCell(plan.name, idx, true),
        createDataCell(plan.ansCode, idx),
        createDataCell(formatCurrency(plan.perCapita), idx),
        createDataCell(formatCurrency(plan.estimatedInvoice), idx),
      ]);
    });

    slide.addTable(tableData, {
      x: 0.3, y: 0.8, w: 9.4,
      colW: [3.5, 2, 2, 1.9],
      border: { pt: 0.5, color: "CCCCCC" },
    });
  }

  // =========================================================================
  // SLIDE 3: VALORES POR FAIXA ETÁRIA - SEM COPARTICIPAÇÃO
  // =========================================================================
  if (data.ageBasedPricingNoCopay && data.ageBasedPricingNoCopay.length > 0) {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    
    slide.addText("Valores por Faixa Etária - SEM Coparticipação", {
      x: 0.3, y: 0.15, w: 9, h: 0.4,
      fontSize: 18, bold: true, color: KLINI_COLORS.TEAL_PRIMARY, fontFace: "Arial",
    });

    const firstRow = data.ageBasedPricingNoCopay[0];
    const planNames = Object.keys(firstRow).filter(k => k !== "ageRange");
    
    const tableData: PptxGenJS.TableRow[] = [];
    
    // Header
    const headerRow: PptxGenJS.TableCell[] = [createHeaderCell("FAIXA ETÁRIA")];
    planNames.forEach(name => {
      headerRow.push(createHeaderCell(name.replace("KLINI ", "K").replace(" EMP ", " ")));
    });
    tableData.push(headerRow);

    // Data rows
    data.ageBasedPricingNoCopay.forEach((row, idx) => {
      const dataRow: PptxGenJS.TableCell[] = [createDataCell(row.ageRange, idx, true)];
      planNames.forEach(name => {
        dataRow.push(createDataCell(formatCurrency(row[name] || 0), idx));
      });
      tableData.push(dataRow);
    });

    const numCols = planNames.length + 1;
    const totalWidth = 9.5;
    const firstColWidth = 0.7;
    const otherColWidth = (totalWidth - firstColWidth) / planNames.length;
    const colWidths = [firstColWidth, ...Array(planNames.length).fill(otherColWidth)];

    slide.addTable(tableData, {
      x: 0.25, y: 0.6, w: totalWidth,
      colW: colWidths,
      border: { pt: 0.5, color: "CCCCCC" },
    });
  }

  // =========================================================================
  // SLIDE 4: PLANOS COM COPARTICIPAÇÃO
  // =========================================================================
  if (data.plansWithCopay && data.plansWithCopay.length > 0) {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    
    slide.addText("Planos com Coparticipação", {
      x: 0.3, y: 0.2, w: 9, h: 0.5,
      fontSize: 22, bold: true, color: KLINI_COLORS.ORANGE, fontFace: "Arial",
    });
    
    slide.addText("ANS - Nº 42.202-9", {
      x: 8.2, y: 0.2, w: 1.5, h: 0.3,
      fontSize: 9, color: KLINI_COLORS.TEXT_GRAY, align: "right",
    });

    const tableData: PptxGenJS.TableRow[] = [];
    
    tableData.push([
      createHeaderCell("PLANO"),
      createHeaderCell("REGISTRO ANS"),
      createHeaderCell("VALOR PER CAPITA"),
      createHeaderCell("FATURA ESTIMADA"),
    ]);

    data.plansWithCopay.forEach((plan, idx) => {
      tableData.push([
        createDataCell(plan.name, idx, true),
        createDataCell(plan.ansCode, idx),
        createDataCell(formatCurrency(plan.perCapita), idx),
        createDataCell(formatCurrency(plan.estimatedInvoice), idx),
      ]);
    });

    slide.addTable(tableData, {
      x: 0.3, y: 0.8, w: 9.4,
      colW: [3.5, 2, 2, 1.9],
      border: { pt: 0.5, color: "CCCCCC" },
    });
  }

  // =========================================================================
  // SLIDE 5: VALORES POR FAIXA ETÁRIA - COM COPARTICIPAÇÃO
  // =========================================================================
  if (data.ageBasedPricingCopay && data.ageBasedPricingCopay.length > 0) {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    
    slide.addText("Valores por Faixa Etária - COM Coparticipação", {
      x: 0.3, y: 0.15, w: 9, h: 0.4,
      fontSize: 18, bold: true, color: KLINI_COLORS.ORANGE, fontFace: "Arial",
    });

    const firstRow = data.ageBasedPricingCopay[0];
    const planNames = Object.keys(firstRow).filter(k => k !== "ageRange");
    
    const tableData: PptxGenJS.TableRow[] = [];
    
    const headerRow: PptxGenJS.TableCell[] = [createHeaderCell("FAIXA ETÁRIA")];
    planNames.forEach(name => {
      headerRow.push(createHeaderCell(name.replace("KLINI ", "K").replace(" EMP ", " ")));
    });
    tableData.push(headerRow);

    data.ageBasedPricingCopay.forEach((row, idx) => {
      const dataRow: PptxGenJS.TableCell[] = [createDataCell(row.ageRange, idx, true)];
      planNames.forEach(name => {
        dataRow.push(createDataCell(formatCurrency(row[name] || 0), idx));
      });
      tableData.push(dataRow);
    });

    const numCols = planNames.length + 1;
    const totalWidth = 9.5;
    const firstColWidth = 0.7;
    const otherColWidth = (totalWidth - firstColWidth) / planNames.length;
    const colWidths = [firstColWidth, ...Array(planNames.length).fill(otherColWidth)];

    slide.addTable(tableData, {
      x: 0.25, y: 0.6, w: totalWidth,
      colW: colWidths,
      border: { pt: 0.5, color: "CCCCCC" },
    });
  }

  // =========================================================================
  // SLIDES PRODUTOS G (se habilitado)
  // =========================================================================
  if (includeProductosG) {
    // SLIDE: PLANOS SEM COPART - PRODUTOS G
    if (data.plansWithoutCopayG && data.plansWithoutCopayG.length > 0) {
      const slide = pptx.addSlide();
      slide.background = { color: "FFFFFF" };
      
      slide.addText("Planos sem Coparticipação - PRODUTOS G", {
        x: 0.3, y: 0.2, w: 9, h: 0.5,
        fontSize: 22, bold: true, color: KLINI_COLORS.TEAL_PRIMARY, fontFace: "Arial",
      });

      const tableData: PptxGenJS.TableRow[] = [];
      tableData.push([
        createHeaderCell("PLANO"),
        createHeaderCell("REGISTRO ANS"),
        createHeaderCell("VALOR PER CAPITA"),
        createHeaderCell("FATURA ESTIMADA"),
      ]);

      data.plansWithoutCopayG.forEach((plan, idx) => {
        tableData.push([
          createDataCell(plan.name, idx, true),
          createDataCell(plan.ansCode, idx),
          createDataCell(formatCurrency(plan.perCapita), idx),
          createDataCell(formatCurrency(plan.estimatedInvoice), idx),
        ]);
      });

      slide.addTable(tableData, {
        x: 0.3, y: 0.8, w: 9.4,
        colW: [3.5, 2, 2, 1.9],
        border: { pt: 0.5, color: "CCCCCC" },
      });
    }

    // SLIDE: FAIXAS ETÁRIAS SEM COPART - PRODUTOS G
    if (data.ageBasedPricingNoCopayG && data.ageBasedPricingNoCopayG.length > 0) {
      const slide = pptx.addSlide();
      slide.background = { color: "FFFFFF" };
      
      slide.addText("Valores por Faixa Etária - SEM Coparticipação (PRODUTOS G)", {
        x: 0.3, y: 0.15, w: 9, h: 0.4,
        fontSize: 16, bold: true, color: KLINI_COLORS.TEAL_PRIMARY, fontFace: "Arial",
      });

      const firstRow = data.ageBasedPricingNoCopayG[0];
      const planNames = Object.keys(firstRow).filter(k => k !== "ageRange");
      
      const tableData: PptxGenJS.TableRow[] = [];
      const headerRow: PptxGenJS.TableCell[] = [createHeaderCell("FAIXA ETÁRIA")];
      planNames.forEach(name => {
        headerRow.push(createHeaderCell(name.replace("KLINI ", "K").replace(" EMP ", " ")));
      });
      tableData.push(headerRow);

      data.ageBasedPricingNoCopayG.forEach((row, idx) => {
        const dataRow: PptxGenJS.TableCell[] = [createDataCell(row.ageRange, idx, true)];
        planNames.forEach(name => {
          dataRow.push(createDataCell(formatCurrency(row[name] || 0), idx));
        });
        tableData.push(dataRow);
      });

      const totalWidth = 9.5;
      const firstColWidth = 0.7;
      const otherColWidth = (totalWidth - firstColWidth) / planNames.length;
      const colWidths = [firstColWidth, ...Array(planNames.length).fill(otherColWidth)];

      slide.addTable(tableData, {
        x: 0.25, y: 0.6, w: totalWidth,
        colW: colWidths,
        border: { pt: 0.5, color: "CCCCCC" },
      });
    }

    // SLIDE: PLANOS COM COPART - PRODUTOS G
    if (data.plansWithCopayG && data.plansWithCopayG.length > 0) {
      const slide = pptx.addSlide();
      slide.background = { color: "FFFFFF" };
      
      slide.addText("Planos com Coparticipação - PRODUTOS G", {
        x: 0.3, y: 0.2, w: 9, h: 0.5,
        fontSize: 22, bold: true, color: KLINI_COLORS.ORANGE, fontFace: "Arial",
      });

      const tableData: PptxGenJS.TableRow[] = [];
      tableData.push([
        createHeaderCell("PLANO"),
        createHeaderCell("REGISTRO ANS"),
        createHeaderCell("VALOR PER CAPITA"),
        createHeaderCell("FATURA ESTIMADA"),
      ]);

      data.plansWithCopayG.forEach((plan, idx) => {
        tableData.push([
          createDataCell(plan.name, idx, true),
          createDataCell(plan.ansCode, idx),
          createDataCell(formatCurrency(plan.perCapita), idx),
          createDataCell(formatCurrency(plan.estimatedInvoice), idx),
        ]);
      });

      slide.addTable(tableData, {
        x: 0.3, y: 0.8, w: 9.4,
        colW: [3.5, 2, 2, 1.9],
        border: { pt: 0.5, color: "CCCCCC" },
      });
    }

    // SLIDE: FAIXAS ETÁRIAS COM COPART - PRODUTOS G
    if (data.ageBasedPricingCopayG && data.ageBasedPricingCopayG.length > 0) {
      const slide = pptx.addSlide();
      slide.background = { color: "FFFFFF" };
      
      slide.addText("Valores por Faixa Etária - COM Coparticipação (PRODUTOS G)", {
        x: 0.3, y: 0.15, w: 9, h: 0.4,
        fontSize: 16, bold: true, color: KLINI_COLORS.ORANGE, fontFace: "Arial",
      });

      const firstRow = data.ageBasedPricingCopayG[0];
      const planNames = Object.keys(firstRow).filter(k => k !== "ageRange");
      
      const tableData: PptxGenJS.TableRow[] = [];
      const headerRow: PptxGenJS.TableCell[] = [createHeaderCell("FAIXA ETÁRIA")];
      planNames.forEach(name => {
        headerRow.push(createHeaderCell(name.replace("KLINI ", "K").replace(" EMP ", " ")));
      });
      tableData.push(headerRow);

      data.ageBasedPricingCopayG.forEach((row, idx) => {
        const dataRow: PptxGenJS.TableCell[] = [createDataCell(row.ageRange, idx, true)];
        planNames.forEach(name => {
          dataRow.push(createDataCell(formatCurrency(row[name] || 0), idx));
        });
        tableData.push(dataRow);
      });

      const totalWidth = 9.5;
      const firstColWidth = 0.7;
      const otherColWidth = (totalWidth - firstColWidth) / planNames.length;
      const colWidths = [firstColWidth, ...Array(planNames.length).fill(otherColWidth)];

      slide.addTable(tableData, {
        x: 0.25, y: 0.6, w: totalWidth,
        colW: colWidths,
        border: { pt: 0.5, color: "CCCCCC" },
      });
    }
  }

  // =========================================================================
  // SALVAR ARQUIVO
  // =========================================================================
  await pptx.writeFile({ fileName: "Proposta_Klini_Saude.pptx" });
}
