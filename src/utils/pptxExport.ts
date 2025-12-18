import pptxgen from "pptxgenjs";

interface ExportData {
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

const formatCurrency = (value: any): string => {
  if (!value) return "R$ 0,00";
  const num = typeof value === "number" ? value : parseFloat(value);
  if (isNaN(num)) return "R$ 0,00";
  return `R$ ${num.toFixed(2).replace(".", ",")}`;
};

export const exportToPPTX = async (data: ExportData, coverImage?: string | null, includeProductosG: boolean = false) => {
  const pptx = new pptxgen();
  pptx.defineLayout({ name: "PORTRAIT_A4", width: 8.26, height: 11.69 });
  pptx.layout = "PORTRAIT_A4";

  const kliniTeal = "1D7874";
  const kliniOrange = "F7931E";
  const lightGray = "F5F5F5";
  const white = "FFFFFF";

  // ==================== SLIDE 1: CAPA ====================
  const slide1 = pptx.addSlide();

  if (coverImage && coverImage.startsWith("data:")) {
    slide1.addImage({
      data: coverImage,
      x: 0,
      y: 0,
      w: 8.26,
      h: 11.69,
      sizing: { type: "cover" },
    });
  } else {
    slide1.background = { color: kliniTeal };
    slide1.addShape(pptx.ShapeType.ellipse, {
      x: 1.5,
      y: 3.0,
      w: 5.0,
      h: 5.0,
      fill: { color: "164E4B", transparency: 30 },
      line: { type: "none" },
    });
    slide1.addText([{ text: "klini", options: { fontSize: 72, bold: true, color: white } }], {
      x: 1.8,
      y: 1.8,
      w: 4.66,
      h: 1.5,
      align: "center",
    });
    slide1.addText([{ text: "saúde", options: { fontSize: 48, bold: false, color: white } }], {
      x: 2.3,
      y: 2.8,
      w: 3.66,
      h: 1.0,
      align: "center",
    });
    slide1.addShape(pptx.ShapeType.rect, {
      x: 0.5,
      y: 6.8,
      w: 7.26,
      h: 1.0,
      fill: { color: white },
      line: { type: "none" },
    });
    slide1.addText("PROPOSTA COMERCIAL", {
      x: 0.5,
      y: 6.85,
      w: 7.26,
      h: 0.9,
      fontSize: 32,
      bold: true,
      color: kliniTeal,
      align: "center",
      valign: "middle",
    });
    const currentDate = new Date().toLocaleDateString("pt-BR", { month: "long", year: "numeric" });
    slide1.addShape(pptx.ShapeType.rect, {
      x: 4.7,
      y: 8.0,
      w: 2.8,
      h: 0.6,
      fill: { color: kliniOrange },
      line: { type: "none" },
    });
    slide1.addText(currentDate.toUpperCase(), {
      x: 4.7,
      y: 8.05,
      w: 2.8,
      h: 0.5,
      fontSize: 16,
      color: white,
      align: "center",
      valign: "middle",
      italic: true,
    });
    slide1.addText("V2.01120251.0", {
      x: 0.3,
      y: 11.35,
      w: 2.0,
      h: 0.25,
      fontSize: 7,
      color: white,
      align: "left",
    });
    slide1.addText("ANS - Nº 42.202-9", {
      x: 6.0,
      y: 11.35,
      w: 2.0,
      h: 0.25,
      fontSize: 7,
      color: white,
      align: "right",
    });
  }

  // ==================== SLIDE 2: DADOS + DEMOGRAFIA ====================
  const slide2 = pptx.addSlide();
  slide2.background = { color: white };

  slide2.addText("Tabela de Preços - ACIMA DE 100 VIDAS", {
    x: 0.5,
    y: 0.3,
    w: 7.26,
    h: 0.4,
    fontSize: 18,
    bold: true,
    color: kliniOrange,
    align: "center",
  });

  slide2.addText("DADOS DA EMPRESA", {
    x: 0.5,
    y: 0.8,
    w: 7.26,
    h: 0.3,
    fontSize: 12,
    bold: true,
    color: kliniTeal,
  });

  let yPos = 1.2;
  const companyInfo = [
    `Razão Social: ${data.companyName}`,
    `Concessionária: ${data.concessionaire}`,
    `Corretor(a): ${data.broker}`,
    `Emissão: ${data.emissionDate} | Validade: ${data.validityDate}`,
  ];

  companyInfo.forEach((info) => {
    slide2.addText(info, {
      x: 0.7,
      y: yPos,
      w: 6.8,
      h: 0.25,
      fontSize: 9,
      color: "333333",
    });
    yPos += 0.35;
  });

  slide2.addText("DEMOGRAFIA", {
    x: 0.5,
    y: 2.5,
    w: 7.26,
    h: 0.3,
    fontSize: 12,
    bold: true,
    color: kliniTeal,
  });

  // Tabela Demografia - COM FILLCOLOR EM CELL
  const demoRows: any[] = [];
  
  // Header
  const headerRow: any[] = [];
  ["FAIXA ETÁRIA", "TITULAR M", "TITULAR F", "DEP. M", "DEP. F", "AGRE. M", "AGRE. F", "TOTAL M", "TOTAL F", "TOTAL", "%"].forEach(label => {
    headerRow.push({
      text: label,
      options: {
        bold: true,
        color: white,
        fill: kliniTeal,
        fontSize: 7,
        border: { pt: 0.5, color: "CCCCCC" }
      }
    });
  });
  demoRows.push(headerRow);

  // Dados com cores alternadas
  data.demographics.forEach((row, idx) => {
    const bgColor = idx % 2 === 0 ? white : lightGray;
    const dataRow: any[] = [
      { text: row.ageRange, options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: String(row.titularM || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: String(row.titularF || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: String(row.dependentM || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: String(row.dependentF || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: String(row.agregadoM || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: String(row.agregadoF || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: String(row.totalM || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: String(row.totalF || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: String(row.total || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: String(row.percentage || "0%"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
    ];
    demoRows.push(dataRow);
  });

  slide2.addTable(demoRows, {
    x: 0.3,
    y: 2.95,
    w: 7.66,
    border: { pt: 0.5, color: "CCCCCC" },
    rowH: 0.25,
  });

  slide2.addText("Esta proposta foi elaborada levando em consideração as informações fornecidas através do formulário de cotação enviado pela Corretora.", {
    x: 0.5,
    y: 10.8,
    w: 7.26,
    h: 0.7,
    fontSize: 6,
    color: "999999",
    align: "center",
    valign: "top",
  });

  // ==================== SLIDE 3: SEM COPARTICIPAÇÃO ====================
  const slide3 = pptx.addSlide();
  slide3.background = { color: white };

  slide3.addText("ANS - Nº 42.202-9", {
    x: 6.5,
    y: 0.2,
    w: 1.5,
    h: 0.3,
    fontSize: 9,
    color: "333333",
    align: "right",
  });

  slide3.addText("Planos sem Coparticipação - ACIMA DE 100 VIDAS", {
    x: 0.5,
    y: 0.5,
    w: 7.26,
    h: 0.4,
    fontSize: 16,
    bold: true,
    color: kliniOrange,
    align: "center",
  });

  const noCopayRows: any[] = [];
  noCopayRows.push([
    { text: "PLANO", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
    { text: "CÓDIGO ANS", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
    { text: "PER CAPITA", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
    { text: "FATURA", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
  ]);

  data.plansWithoutCopay.forEach((plan, idx) => {
    const bgColor = idx % 2 === 0 ? white : lightGray;
    noCopayRows.push([
      { text: plan.name, options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: String(plan.ansCode), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: formatCurrency(plan.perCapita), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: formatCurrency(plan.estimatedInvoice), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
    ]);
  });

  slide3.addTable(noCopayRows, {
    x: 0.5,
    y: 1.1,
    w: 7.26,
    border: { pt: 0.5, color: "CCCCCC" },
    align: "center",
  });

  slide3.addText("Valores por Faixa Etária - SEM Coparticipação", {
    x: 0.5,
    y: 3.2,
    w: 7.26,
    h: 0.35,
    fontSize: 11,
    bold: true,
    color: kliniOrange,
    align: "center",
  });

  if (data.ageBasedPricingNoCopay && data.ageBasedPricingNoCopay.length > 0) {
    const planColumns = Object.keys(data.ageBasedPricingNoCopay[0]).filter((key) => key !== "ageRange");
    const ageRowsNoCopay: any[] = [];
    
    ageRowsNoCopay.push([
      { text: "FAIXA ETÁRIA", options: { bold: true, color: white, fill: kliniTeal, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } } },
      ...planColumns.map((planName) => ({
        text: planName.substring(0, 12),
        options: { bold: true, color: white, fill: kliniTeal, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } },
      })),
    ]);

    data.ageBasedPricingNoCopay.forEach((row, idx) => {
      const bgColor = idx % 2 === 0 ? white : lightGray;
      ageRowsNoCopay.push([
        { text: row.ageRange, options: { color: "333333", fill: bgColor, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } } },
        ...planColumns.map((col) => ({
          text: row[col] ? formatCurrency(row[col]) : "-",
          options: { color: "333333", fill: bgColor, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } },
        })),
      ]);
    });

    const colWidth = (7.26 - 0.8) / planColumns.length;
    slide3.addTable(ageRowsNoCopay, {
      x: 0.5,
      y: 3.7,
      w: 7.26,
      border: { pt: 0.5, color: "CCCCCC" },
      align: "center",
      colW: [0.8, ...Array(planColumns.length).fill(colWidth)],
    });
  }

  slide3.addText("Esta proposta foi elaborada levando em consideração as informações fornecidas através do formulário de cotação enviado pela Corretora.", {
    x: 0.5,
    y: 10.8,
    w: 7.26,
    h: 0.7,
    fontSize: 6,
    color: "999999",
    align: "center",
    valign: "top",
  });

  // ==================== SLIDE 4: COM COPARTICIPAÇÃO ====================
  const slide4 = pptx.addSlide();
  slide4.background = { color: white };

  slide4.addText("ANS - Nº 42.202-9", {
    x: 6.5,
    y: 0.2,
    w: 1.5,
    h: 0.3,
    fontSize: 9,
    color: "333333",
    align: "right",
  });

  slide4.addText("Planos com Coparticipação - ACIMA DE 100 VIDAS", {
    x: 0.5,
    y: 0.5,
    w: 7.26,
    h: 0.4,
    fontSize: 16,
    bold: true,
    color: kliniOrange,
    align: "center",
  });

  const copayRows: any[] = [];
  copayRows.push([
    { text: "PLANO", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
    { text: "CÓDIGO ANS", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
    { text: "PER CAPITA", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
    { text: "FATURA", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
  ]);

  data.plansWithCopay.forEach((plan, idx) => {
    const bgColor = idx % 2 === 0 ? white : lightGray;
    copayRows.push([
      { text: plan.name, options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: String(plan.ansCode), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: formatCurrency(plan.perCapita), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: formatCurrency(plan.estimatedInvoice), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
    ]);
  });

  slide4.addTable(copayRows, {
    x: 0.5,
    y: 1.1,
    w: 7.26,
    border: { pt: 0.5, color: "CCCCCC" },
    align: "center",
  });

  slide4.addText("Valores por Faixa Etária - COM Coparticipação", {
    x: 0.5,
    y: 3.2,
    w: 7.26,
    h: 0.35,
    fontSize: 11,
    bold: true,
    color: kliniOrange,
    align: "center",
  });

  if (data.ageBasedPricingCopay && data.ageBasedPricingCopay.length > 0) {
    const planColumns = Object.keys(data.ageBasedPricingCopay[0]).filter((key) => key !== "ageRange");
    const ageRowsCopay: any[] = [];
    
    ageRowsCopay.push([
      { text: "FAIXA ETÁRIA", options: { bold: true, color: white, fill: kliniTeal, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } } },
      ...planColumns.map((planName) => ({
        text: planName.substring(0, 12),
        options: { bold: true, color: white, fill: kliniTeal, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } },
      })),
    ]);

    data.ageBasedPricingCopay.forEach((row, idx) => {
      const bgColor = idx % 2 === 0 ? white : lightGray;
      ageRowsCopay.push([
        { text: row.ageRange, options: { color: "333333", fill: bgColor, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } } },
        ...planColumns.map((col) => ({
          text: row[col] ? formatCurrency(row[col]) : "-",
          options: { color: "333333", fill: bgColor, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } },
        })),
      ]);
    });

    const colWidth = (7.26 - 0.8) / planColumns.length;
    slide4.addTable(ageRowsCopay, {
      x: 0.5,
      y: 3.7,
      w: 7.26,
      border: { pt: 0.5, color: "CCCCCC" },
      align: "center",
      colW: [0.8, ...Array(planColumns.length).fill(colWidth)],
    });
  }

  slide4.addText("Esta proposta foi elaborada levando em consideração as informações fornecidas através do formulário de cotação enviado pela Corretora.", {
    x: 0.5,
    y: 10.8,
    w: 7.26,
    h: 0.7,
    fontSize: 6,
    color: "999999",
    align: "center",
    valign: "top",
  });

  // ==================== SLIDES 5-7: PRODUTOS G ====================
  if (includeProductosG && data.demographicsG && data.demographicsG.length > 0) {
    // SLIDE 5
    const slide5 = pptx.addSlide();
    slide5.background = { color: white };

    slide5.addText("Tabela de Preços - PRODUTOS G", {
      x: 0.5,
      y: 0.3,
      w: 7.26,
      h: 0.4,
      fontSize: 18,
      bold: true,
      color: kliniOrange,
      align: "center",
    });

    slide5.addText("DEMOGRAFIA - PRODUTOS G", {
      x: 0.5,
      y: 0.8,
      w: 7.26,
      h: 0.3,
      fontSize: 12,
      bold: true,
      color: kliniTeal,
    });

    const demGRows: any[] = [];
    demGRows.push([
      { text: "FAIXA ETÁRIA", options: { bold: true, color: white, fill: kliniTeal, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: "TITULAR M", options: { bold: true, color: white, fill: kliniTeal, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: "TITULAR F", options: { bold: true, color: white, fill: kliniTeal, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: "DEP. M", options: { bold: true, color: white, fill: kliniTeal, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: "DEP. F", options: { bold: true, color: white, fill: kliniTeal, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: "AGRE. M", options: { bold: true, color: white, fill: kliniTeal, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: "AGRE. F", options: { bold: true, color: white, fill: kliniTeal, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: "TOTAL M", options: { bold: true, color: white, fill: kliniTeal, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: "TOTAL F", options: { bold: true, color: white, fill: kliniTeal, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: "TOTAL", options: { bold: true, color: white, fill: kliniTeal, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      { text: "%", options: { bold: true, color: white, fill: kliniTeal, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
    ]);

    data.demographicsG.forEach((row, idx) => {
      const bgColor = idx % 2 === 0 ? white : lightGray;
      demGRows.push([
        { text: row.ageRange, options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: String(row.titularM || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: String(row.titularF || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: String(row.dependentM || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: String(row.dependentF || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: String(row.agregadoM || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: String(row.agregadoF || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: String(row.totalM || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: String(row.totalF || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: String(row.total || "0"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: String(row.percentage || "0%"), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
      ]);
    });

    slide5.addTable(demGRows, {
      x: 0.3,
      y: 1.25,
      w: 7.66,
      border: { pt: 0.5, color: "CCCCCC" },
      rowH: 0.25,
    });

    slide5.addText("Esta proposta foi elaborada levando em consideração as informações fornecidas através do formulário de cotação enviado pela Corretora.", {
      x: 0.5,
      y: 10.8,
      w: 7.26,
      h: 0.7,
      fontSize: 6,
      color: "999999",
      align: "center",
      valign: "top",
    });

    // SLIDE 6
    const slide6 = pptx.addSlide();
    slide6.background = { color: white };

    slide6.addText("ANS - Nº 42.202-9", {
      x: 6.5,
      y: 0.2,
      w: 1.5,
      h: 0.3,
      fontSize: 9,
      color: "333333",
      align: "right",
    });

    slide6.addText("Planos sem Coparticipação - PRODUTOS G", {
      x: 0.5,
      y: 0.5,
      w: 7.26,
      h: 0.4,
      fontSize: 16,
      bold: true,
      color: kliniOrange,
      align: "center",
    });

    if (data.plansWithoutCopayG && data.plansWithoutCopayG.length > 0) {
      const noCopayGRows: any[] = [];
      noCopayGRows.push([
        { text: "PLANO", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: "CÓDIGO ANS", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: "PER CAPITA", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: "FATURA", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
      ]);

      data.plansWithoutCopayG.forEach((plan, idx) => {
        const bgColor = idx % 2 === 0 ? white : lightGray;
        noCopayGRows.push([
          { text: plan.name, options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
          { text: String(plan.ansCode), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
          { text: formatCurrency(plan.perCapita), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
          { text: formatCurrency(plan.estimatedInvoice), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
        ]);
      });

      slide6.addTable(noCopayGRows, {
        x: 0.5,
        y: 1.1,
        w: 7.26,
        border: { pt: 0.5, color: "CCCCCC" },
        align: "center",
      });

      slide6.addText("Valores por Faixa Etária - SEM Coparticipação", {
        x: 0.5,
        y: 3.2,
        w: 7.26,
        h: 0.35,
        fontSize: 11,
        bold: true,
        color: kliniOrange,
        align: "center",
      });

      if (data.ageBasedPricingNoCopayG && data.ageBasedPricingNoCopayG.length > 0) {
        const planColumns = Object.keys(data.ageBasedPricingNoCopayG[0]).filter((key) => key !== "ageRange");
        const ageRowsNoCopayG: any[] = [];
        ageRowsNoCopayG.push([
          { text: "FAIXA ETÁRIA", options: { bold: true, color: white, fill: kliniTeal, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } } },
          ...planColumns.map((planName) => ({
            text: planName.substring(0, 12),
            options: { bold: true, color: white, fill: kliniTeal, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } },
          })),
        ]);

        data.ageBasedPricingNoCopayG.forEach((row, idx) => {
          const bgColor = idx % 2 === 0 ? white : lightGray;
          ageRowsNoCopayG.push([
            { text: row.ageRange, options: { color: "333333", fill: bgColor, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } } },
            ...planColumns.map((col) => ({
              text: row[col] ? formatCurrency(row[col]) : "-",
              options: { color: "333333", fill: bgColor, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } },
            })),
          ]);
        });

        const colWidth = (7.26 - 0.8) / planColumns.length;
        slide6.addTable(ageRowsNoCopayG, {
          x: 0.5,
          y: 3.7,
          w: 7.26,
          border: { pt: 0.5, color: "CCCCCC" },
          align: "center",
          colW: [0.8, ...Array(planColumns.length).fill(colWidth)],
        });
      }
    }

    slide6.addText("Esta proposta foi elaborada levando em consideração as informações fornecidas através do formulário de cotação enviado pela Corretora.", {
      x: 0.5,
      y: 10.8,
      w: 7.26,
      h: 0.7,
      fontSize: 6,
      color: "999999",
      align: "center",
      valign: "top",
    });

    // SLIDE 7
    const slide7 = pptx.addSlide();
    slide7.background = { color: white };

    slide7.addText("ANS - Nº 42.202-9", {
      x: 6.5,
      y: 0.2,
      w: 1.5,
      h: 0.3,
      fontSize: 9,
      color: "333333",
      align: "right",
    });

    slide7.addText("Planos com Coparticipação - PRODUTOS G", {
      x: 0.5,
      y: 0.5,
      w: 7.26,
      h: 0.4,
      fontSize: 16,
      bold: true,
      color: kliniOrange,
      align: "center",
    });

    if (data.plansWithCopayG && data.plansWithCopayG.length > 0) {
      const copayGRows: any[] = [];
      copayGRows.push([
        { text: "PLANO", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: "CÓDIGO ANS", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: "PER CAPITA", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
        { text: "FATURA", options: { bold: true, color: white, fill: kliniTeal, fontSize: 8, border: { pt: 0.5, color: "CCCCCC" } } },
      ]);

      data.plansWithCopayG.forEach((plan, idx) => {
        const bgColor = idx % 2 === 0 ? white : lightGray;
        copayGRows.push([
          { text: plan.name, options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
          { text: String(plan.ansCode), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
          { text: formatCurrency(plan.perCapita), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
          { text: formatCurrency(plan.estimatedInvoice), options: { color: "333333", fill: bgColor, fontSize: 7, border: { pt: 0.5, color: "CCCCCC" } } },
        ]);
      });

      slide7.addTable(copayGRows, {
        x: 0.5,
        y: 1.1,
        w: 7.26,
        border: { pt: 0.5, color: "CCCCCC" },
        align: "center",
      });

      slide7.addText("Valores por Faixa Etária - COM Coparticipação", {
        x: 0.5,
        y: 3.2,
        w: 7.26,
        h: 0.35,
        fontSize: 11,
        bold: true,
        color: kliniOrange,
        align: "center",
      });

      if (data.ageBasedPricingCopayG && data.ageBasedPricingCopayG.length > 0) {
        const planColumns = Object.keys(data.ageBasedPricingCopayG[0]).filter((key) => key !== "ageRange");
        const ageRowsCopayG: any[] = [];
        ageRowsCopayG.push([
          { text: "FAIXA ETÁRIA", options: { bold: true, color: white, fill: kliniTeal, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } } },
          ...planColumns.map((planName) => ({
            text: planName.substring(0, 12),
            options: { bold: true, color: white, fill: kliniTeal, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } },
          })),
        ]);

        data.ageBasedPricingCopayG.forEach((row, idx) => {
          const bgColor = idx % 2 === 0 ? white : lightGray;
          ageRowsCopayG.push([
            { text: row.ageRange, options: { color: "333333", fill: bgColor, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } } },
            ...planColumns.map((col) => ({
              text: row[col] ? formatCurrency(row[col]) : "-",
              options: { color: "333333", fill: bgColor, fontSize: 6, border: { pt: 0.5, color: "CCCCCC" } },
            })),
          ]);
        });

        const colWidth = (7.26 - 0.8) / planColumns.length;
        slide7.addTable(ageRowsCopayG, {
          x: 0.5,
          y: 3.7,
          w: 7.26,
          border: { pt: 0.5, color: "CCCCCC" },
          align: "center",
          colW: [0.8, ...Array(planColumns.length).fill(colWidth)],
        });
      }
    }

    slide7.addText("Esta proposta foi elaborada levando em consideração as informações fornecidas através do formulário de cotação enviado pela Corretora.", {
      x: 0.5,
      y: 10.8,
      w: 7.26,
      h: 0.7,
      fontSize: 6,
      color: "999999",
      align: "center",
      valign: "top",
    });
  }

  await pptx.writeFile({ fileName: `Proposta_${data.companyName || "Klini_Saude"}.pptx` });
};
