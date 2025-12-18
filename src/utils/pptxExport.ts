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
  // PRODUTOS G
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

  // ============================================================================
  // PALETA DE CORES KLINI - CORRIGIDA
  // ============================================================================
  const kliniTeal = "1D7874";           // Verde teal escuro (cabeçalhos)
  const kliniOrange = "F7931E";         // Laranja
  const rowWhite = "FFFFFF";            // Linha branca
  const rowGreenLight = "E8F5F3";       // Linha verde menta claro (CORRIGIDO!)
  const textDark = "1D3D3A";            // Texto verde escuro

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
    slide1.addText([{ text: "klini", options: { fontSize: 72, bold: true, color: "FFFFFF" } }], {
      x: 1.8,
      y: 1.8,
      w: 4.66,
      h: 1.5,
      align: "center",
    });
    slide1.addText([{ text: "saúde", options: { fontSize: 48, bold: false, color: "FFFFFF" } }], {
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
      fill: { color: "FFFFFF" },
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
      color: "FFFFFF",
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
      color: "FFFFFF",
      align: "left",
    });
    slide1.addText("ANS - Nº 42.202-9", {
      x: 6.0,
      y: 11.35,
      w: 2.0,
      h: 0.25,
      fontSize: 7,
      color: "FFFFFF",
      align: "right",
    });
  }

  // ==================== SLIDE 2: DADOS + DEMOGRAFIA ====================
  const slide2 = pptx.addSlide();
  slide2.background = { color: "FFFFFF" };

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
      color: textDark,
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
    const bgColor = idx % 2 === 0 ? rowWhite : rowGreenLight;
    demoRows.push([
      { text: row.ageRange, options: { bold: true, color: textDark, fill: { color: rowGreenLight }, fontSize: 7 } },
      { text: String(row.titularM || "0"), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.titularF || "0"), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.dependentM || "0"), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.dependentF || "0"), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.agregadoM || "0"), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.agregadoF || "0"), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.totalM || "0"), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.totalF || "0"), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.total || "0"), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
      { text: String(row.percentage || "0%"), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
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
  slide3.background = { color: "FFFFFF" };

  slide3.addText("ANS - Nº 42.202-9", {
    x: 6.5,
    y: 0.2,
    w: 1.5,
    h: 0.3,
    fontSize: 9,
    color: textDark,
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
    const bgColor = idx % 2 === 0 ? rowWhite : rowGreenLight;
    noCopayRows.push([
      { text: plan.name, options: { bold: true, color: textDark, fill: { color: rowGreenLight }, fontSize: 7 } },
      { text: String(plan.ansCode), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
      { text: formatCurrency(plan.perCapita), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
      { text: formatCurrency(plan.estimatedInvoice), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
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
      const bgColor = idx % 2 === 0 ? rowWhite : rowGreenLight;
      ageRowsNoCopay.push([
        { text: row.ageRange, options: { bold: true, color: textDark, fill: { color: rowGreenLight }, fontSize: 6 } },
        ...planColumns.map((col) => ({
          text: row[col] ? formatCurrency(row[col]) : "-",
          options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 6 },
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
  slide4.background = { color: "FFFFFF" };

  slide4.addText("ANS - Nº 42.202-9", {
    x: 6.5,
    y: 0.2,
    w: 1.5,
    h: 0.3,
    fontSize: 9,
    color: textDark,
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
    const bgColor = idx % 2 === 0 ? rowWhite : rowGreenLight;
    copayRows.push([
      { text: plan.name, options: { bold: true, color: textDark, fill: { color: rowGreenLight }, fontSize: 7 } },
      { text: String(plan.ansCode), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
      { text: formatCurrency(plan.perCapita), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
      { text: formatCurrency(plan.estimatedInvoice), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
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
      const bgColor = idx % 2 === 0 ? rowWhite : rowGreenLight;
      ageRowsCopay.push([
        { text: row.ageRange, options: { bold: true, color: textDark, fill: { color: rowGreenLight }, fontSize: 6 } },
        ...planColumns.map((col) => ({
          text: row[col] ? formatCurrency(row[col]) : "-",
          options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 6 },
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

  // ==================== SLIDES PRODUTOS G (SE HABILITADO) ====================
  if (includeProductosG) {
    // SLIDE 5: PRODUTOS G - SEM COPARTICIPAÇÃO
    if (data.plansWithoutCopayG && data.plansWithoutCopayG.length > 0) {
      const slideG1 = pptx.addSlide();
      slideG1.background = { color: "FFFFFF" };

      slideG1.addText("ANS - Nº 42.202-9", {
        x: 6.5,
        y: 0.2,
        w: 1.5,
        h: 0.3,
        fontSize: 9,
        color: textDark,
        align: "right",
      });

      slideG1.addText("PRODUTOS G - Planos sem Coparticipação", {
        x: 0.5,
        y: 0.5,
        w: 7.26,
        h: 0.4,
        fontSize: 14,
        bold: true,
        color: kliniOrange,
        align: "center",
      });

      const noCopayRowsG: any[] = [
        [
          { text: "PLANO", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
          { text: "CÓDIGO ANS", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
          { text: "PER CAPITA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
          { text: "FATURA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
        ],
      ];

      data.plansWithoutCopayG.forEach((plan, idx) => {
        const bgColor = idx % 2 === 0 ? rowWhite : rowGreenLight;
        noCopayRowsG.push([
          { text: plan.name, options: { bold: true, color: textDark, fill: { color: rowGreenLight }, fontSize: 7 } },
          { text: String(plan.ansCode), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
          { text: formatCurrency(plan.perCapita), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
          { text: formatCurrency(plan.estimatedInvoice), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
        ]);
      });

      slideG1.addTable(noCopayRowsG, {
        x: 0.5,
        y: 1.1,
        w: 7.26,
        border: { pt: 0.5, color: "CCCCCC" },
        align: "center",
      });

      // Faixas Etárias SEM COPAY - PRODUTOS G
      if (data.ageBasedPricingNoCopayG && data.ageBasedPricingNoCopayG.length > 0) {
        slideG1.addText("Valores por Faixa Etária - SEM Coparticipação", {
          x: 0.5,
          y: 3.0,
          w: 7.26,
          h: 0.35,
          fontSize: 11,
          bold: true,
          color: kliniOrange,
          align: "center",
        });

        const planColumnsG = Object.keys(data.ageBasedPricingNoCopayG[0]).filter((key) => key !== "ageRange");
        const ageRowsNoCopayG: any[] = [
          [
            { text: "FAIXA ETÁRIA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 6 } },
            ...planColumnsG.map((planName) => ({
              text: planName.substring(0, 12),
              options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 6 },
            })),
          ],
        ];

        data.ageBasedPricingNoCopayG.forEach((row, idx) => {
          const bgColor = idx % 2 === 0 ? rowWhite : rowGreenLight;
          ageRowsNoCopayG.push([
            { text: row.ageRange, options: { bold: true, color: textDark, fill: { color: rowGreenLight }, fontSize: 6 } },
            ...planColumnsG.map((col) => ({
              text: row[col] ? formatCurrency(row[col]) : "-",
              options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 6 },
            })),
          ]);
        });

        const colWidthG = (7.26 - 0.8) / planColumnsG.length;
        slideG1.addTable(ageRowsNoCopayG, {
          x: 0.5,
          y: 3.5,
          w: 7.26,
          border: { pt: 0.5, color: "CCCCCC" },
          align: "center",
          colW: [0.8, ...Array(planColumnsG.length).fill(colWidthG)],
        });
      }

      slideG1.addText("Esta proposta foi elaborada levando em consideração as informações fornecidas através do formulário de cotação enviado pela Corretora.", {
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

    // SLIDE 6: PRODUTOS G - COM COPARTICIPAÇÃO
    if (data.plansWithCopayG && data.plansWithCopayG.length > 0) {
      const slideG2 = pptx.addSlide();
      slideG2.background = { color: "FFFFFF" };

      slideG2.addText("ANS - Nº 42.202-9", {
        x: 6.5,
        y: 0.2,
        w: 1.5,
        h: 0.3,
        fontSize: 9,
        color: textDark,
        align: "right",
      });

      slideG2.addText("PRODUTOS G - Planos com Coparticipação", {
        x: 0.5,
        y: 0.5,
        w: 7.26,
        h: 0.4,
        fontSize: 14,
        bold: true,
        color: kliniOrange,
        align: "center",
      });

      const copayRowsG: any[] = [
        [
          { text: "PLANO", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
          { text: "CÓDIGO ANS", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
          { text: "PER CAPITA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
          { text: "FATURA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 8 } },
        ],
      ];

      data.plansWithCopayG.forEach((plan, idx) => {
        const bgColor = idx % 2 === 0 ? rowWhite : rowGreenLight;
        copayRowsG.push([
          { text: plan.name, options: { bold: true, color: textDark, fill: { color: rowGreenLight }, fontSize: 7 } },
          { text: String(plan.ansCode), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
          { text: formatCurrency(plan.perCapita), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
          { text: formatCurrency(plan.estimatedInvoice), options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 7 } },
        ]);
      });

      slideG2.addTable(copayRowsG, {
        x: 0.5,
        y: 1.1,
        w: 7.26,
        border: { pt: 0.5, color: "CCCCCC" },
        align: "center",
      });

      // Faixas Etárias COM COPAY - PRODUTOS G
      if (data.ageBasedPricingCopayG && data.ageBasedPricingCopayG.length > 0) {
        slideG2.addText("Valores por Faixa Etária - COM Coparticipação", {
          x: 0.5,
          y: 3.0,
          w: 7.26,
          h: 0.35,
          fontSize: 11,
          bold: true,
          color: kliniOrange,
          align: "center",
        });

        const planColumnsG = Object.keys(data.ageBasedPricingCopayG[0]).filter((key) => key !== "ageRange");
        const ageRowsCopayG: any[] = [
          [
            { text: "FAIXA ETÁRIA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 6 } },
            ...planColumnsG.map((planName) => ({
              text: planName.substring(0, 12),
              options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, fontSize: 6 },
            })),
          ],
        ];

        data.ageBasedPricingCopayG.forEach((row, idx) => {
          const bgColor = idx % 2 === 0 ? rowWhite : rowGreenLight;
          ageRowsCopayG.push([
            { text: row.ageRange, options: { bold: true, color: textDark, fill: { color: rowGreenLight }, fontSize: 6 } },
            ...planColumnsG.map((col) => ({
              text: row[col] ? formatCurrency(row[col]) : "-",
              options: { bold: false, color: textDark, fill: { color: bgColor }, fontSize: 6 },
            })),
          ]);
        });

        const colWidthG = (7.26 - 0.8) / planColumnsG.length;
        slideG2.addTable(ageRowsCopayG, {
          x: 0.5,
          y: 3.5,
          w: 7.26,
          border: { pt: 0.5, color: "CCCCCC" },
          align: "center",
          colW: [0.8, ...Array(planColumnsG.length).fill(colWidthG)],
        });
      }

      slideG2.addText("Esta proposta foi elaborada levando em consideração as informações fornecidas através do formulário de cotação enviado pela Corretora.", {
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
  }

  // Salvar arquivo
  await pptx.writeFile({ fileName: `Proposta_${data.companyName || "Klini_Saude"}.pptx` });
};
