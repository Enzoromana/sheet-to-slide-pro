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
}

const formatCurrency = (value: any): string => {
  if (!value) return "R$ 0,00";
  const num = typeof value === "number" ? value : parseFloat(value);
  if (isNaN(num)) return "R$ 0,00";
  return `R$ ${num.toFixed(2).replace(".", ",")}`;
};

export const exportToPPTX = async (data: ExportData, coverImage?: string | null) => {
  const pptx = new pptxgen();
  pptx.defineLayout({ name: "PORTRAIT_A4", width: 8.26, height: 11.69 });
  pptx.layout = "PORTRAIT_A4";

  const kliniTeal = "1D7874";
  const kliniOrange = "F7931E";
  const lightGray = "F5F5F5";

  // ==================== SLIDE 1: CAPA ====================
  const slide1 = pptx.addSlide();

  // Verificar se tem imagem de capa válida em base64
  const hasValidCover = coverImage && typeof coverImage === 'string' && coverImage.startsWith("data:");

  if (hasValidCover) {
    // Usar imagem de capa customizada
    slide1.addImage({
      data: coverImage,
      x: 0,
      y: 0,
      w: 8.26,
      h: 11.69,
      sizing: { type: "cover", w: 8.26, h: 11.69 },
    });
  } else {
    // CAPA PADRÃO - Sempre exibida quando não há imagem
    
    // Fundo verde Klini
    slide1.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: 8.26,
      h: 11.69,
      fill: { color: kliniTeal },
      line: { type: "none" },
    });

    // Círculo decorativo
    slide1.addShape(pptx.ShapeType.ellipse, {
      x: 1.5,
      y: 3.0,
      w: 5.0,
      h: 5.0,
      fill: { color: "164E4B", transparency: 30 },
      line: { type: "none" },
    });

    // Logo "klini"
    slide1.addText("klini", {
      x: 1.8,
      y: 1.8,
      w: 4.66,
      h: 1.5,
      fontSize: 72,
      bold: true,
      color: "FFFFFF",
      align: "center",
      fontFace: "Arial",
    });

    // Subtítulo "saúde"
    slide1.addText("saúde", {
      x: 2.3,
      y: 2.8,
      w: 3.66,
      h: 1.0,
      fontSize: 48,
      bold: false,
      color: "FFFFFF",
      align: "center",
      fontFace: "Arial",
    });

    // Barra branca para "PROPOSTA COMERCIAL"
    slide1.addShape(pptx.ShapeType.rect, {
      x: 0.5,
      y: 6.8,
      w: 7.26,
      h: 1.0,
      fill: { color: "FFFFFF" },
      line: { type: "none" },
    });

    // Texto "PROPOSTA COMERCIAL"
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
      fontFace: "Arial",
    });

    // Data atual
    const currentDate = new Date().toLocaleDateString("pt-BR", { month: "long", year: "numeric" });
    
    // Barra laranja com a data
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
      color: "FFFFFF",
      align: "center",
      valign: "middle",
      italic: true,
      fontFace: "Arial",
    });

    // Versão (canto inferior esquerdo)
    slide1.addText("V2.01120251.0", {
      x: 0.3,
      y: 11.35,
      w: 2.0,
      h: 0.25,
      fontSize: 7,
      color: "FFFFFF",
      align: "left",
      fontFace: "Arial",
    });

    // ANS (canto inferior direito)
    slide1.addText("ANS - Nº 42.202-9", {
      x: 6.0,
      y: 11.35,
      w: 2.0,
      h: 0.25,
      fontSize: 7,
      color: "FFFFFF",
      align: "right",
      fontFace: "Arial",
    });
  }

  // ==================== SLIDE 2: DADOS + DEMOGRAFIA ====================
  const slide2 = pptx.addSlide();
  
  // Fundo branco
  slide2.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 8.26,
    h: 11.69,
    fill: { color: "FFFFFF" },
    line: { type: "none" },
  });

  slide2.addText("Tabela de Preços", {
    x: 0.5,
    y: 0.3,
    w: 7.26,
    h: 0.4,
    fontSize: 20,
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

  // Tabela Demografia
  const demoRows: any[] = [
    [
      { text: "FAIXA ETÁRIA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 7 } },
      { text: "TITULAR M", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 7 } },
      { text: "TITULAR F", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 7 } },
      { text: "DEP. M", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 7 } },
      { text: "DEP. F", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 7 } },
      { text: "AGRE. M", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 7 } },
      { text: "AGRE. F", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 7 } },
      { text: "TOTAL M", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 7 } },
      { text: "TOTAL F", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 7 } },
      { text: "TOTAL", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 7 } },
      { text: "%", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 7 } },
    ],
  ];

  data.demographics.forEach((row, idx) => {
    const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
    demoRows.push([
      { text: row.ageRange, options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.titularM || "0"), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.titularF || "0"), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.dependentM || "0"), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.dependentF || "0"), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.agregadoM || "0"), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.agregadoF || "0"), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.totalM || "0"), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.totalF || "0"), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.total || "0"), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.percentage || "0%"), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
    ]);
  });

  slide2.addTable(demoRows, {
    x: 0.3,
    y: 2.95,
    w: 7.66,
    border: { pt: 0.5, color: "CCCCCC" },
    fontSize: 7,
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
  
  slide3.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 8.26,
    h: 11.69,
    fill: { color: "FFFFFF" },
    line: { type: "none" },
  });

  slide3.addText("ANS - Nº 42.202-9", {
    x: 6.5,
    y: 0.2,
    w: 1.5,
    h: 0.3,
    fontSize: 9,
    color: "333333",
    align: "right",
  });

  slide3.addText("Planos sem Coparticipação", {
    x: 0.5,
    y: 0.5,
    w: 7.26,
    h: 0.4,
    fontSize: 16,
    bold: true,
    color: kliniOrange,
    align: "center",
  });

  // Tabela Planos SEM COPAY
  const noCopayRows: any[] = [
    [
      { text: "PLANO", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
      { text: "CÓDIGO ANS", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
      { text: "PER CAPITA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
      { text: "FATURA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
    ],
  ];

  data.plansWithoutCopay.forEach((plan, idx) => {
    const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
    noCopayRows.push([
      { text: plan.name, options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: String(plan.ansCode), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: formatCurrency(plan.perCapita), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: formatCurrency(plan.estimatedInvoice), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
    ]);
  });

  slide3.addTable(noCopayRows, {
    x: 0.5,
    y: 1.1,
    w: 7.26,
    border: { pt: 0.5, color: "CCCCCC" },
    align: "center",
  });

  // Valores por Faixa Etária SEM COPAY
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
    const ageRowsNoCopay: any[] = [
      [
        { text: "FAIXA ETÁRIA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 6 } },
        ...planColumns.map((planName) => ({
          text: planName.substring(0, 12),
          options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 6 },
        })),
      ],
    ];

    data.ageBasedPricingNoCopay.forEach((row, idx) => {
      const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
      ageRowsNoCopay.push([
        { text: row.ageRange, options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 6 } },
        ...planColumns.map((col) => ({
          text: row[col] ? formatCurrency(row[col]) : "-",
          options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 6 },
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
  
  slide4.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 8.26,
    h: 11.69,
    fill: { color: "FFFFFF" },
    line: { type: "none" },
  });

  slide4.addText("ANS - Nº 42.202-9", {
    x: 6.5,
    y: 0.2,
    w: 1.5,
    h: 0.3,
    fontSize: 9,
    color: "333333",
    align: "right",
  });

  slide4.addText("Planos com Coparticipação", {
    x: 0.5,
    y: 0.5,
    w: 7.26,
    h: 0.4,
    fontSize: 16,
    bold: true,
    color: kliniOrange,
    align: "center",
  });

  // Tabela Planos COM COPAY
  const copayRows: any[] = [
    [
      { text: "PLANO", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
      { text: "CÓDIGO ANS", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
      { text: "PER CAPITA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
      { text: "FATURA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
    ],
  ];

  data.plansWithCopay.forEach((plan, idx) => {
    const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
    copayRows.push([
      { text: plan.name, options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: String(plan.ansCode), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: formatCurrency(plan.perCapita), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
      { text: formatCurrency(plan.estimatedInvoice), options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 7 } },
    ]);
  });

  slide4.addTable(copayRows, {
    x: 0.5,
    y: 1.1,
    w: 7.26,
    border: { pt: 0.5, color: "CCCCCC" },
    align: "center",
  });

  // Valores por Faixa Etária COM COPAY
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
    const ageRowsCopay: any[] = [
      [
        { text: "FAIXA ETÁRIA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 6 } },
        ...planColumns.map((planName) => ({
          text: planName.substring(0, 12),
          options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 6 },
        })),
      ],
    ];

    data.ageBasedPricingCopay.forEach((row, idx) => {
      const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
      ageRowsCopay.push([
        { text: row.ageRange, options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 6 } },
        ...planColumns.map((col) => ({
          text: row[col] ? formatCurrency(row[col]) : "-",
          options: { bold: false, color: "333333", fill: { color: bgColor }, fontSize: 6 },
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

  // Salvar arquivo
  await pptx.writeFile({ fileName: `Proposta_${data.companyName || "Klini_Saude"}.pptx` });
};
