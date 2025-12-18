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

  // PALETA KLINI CORRETA
  const kliniTeal = "1D7874";       // Verde escuro (header)
  const kliniTealLight = "B8D4D3";  // Teal claro (linhas alternadas)
  const kliniOrange = "F7931E";
  const white = "FFFFFF";

  // ==================== SLIDE 1: CAPA ====================
  const slide1 = pptx.addSlide();
  slide1.background = { color: kliniTeal };
  
  slide1.addText("klini", {
    x: 2,
    y: 4,
    w: 4,
    h: 1,
    fontSize: 72,
    bold: true,
    color: white,
    align: "center",
  });

  slide1.addText("saúde", {
    x: 2,
    y: 5,
    w: 4,
    h: 1,
    fontSize: 48,
    bold: false,
    color: white,
    align: "center",
  });

  // ==================== SLIDE 2: DEMOGRAFIA ====================
  const slide2 = pptx.addSlide();
  slide2.background = { color: white };

  slide2.addText("DEMOGRAFIA", {
    x: 0.5,
    y: 0.3,
    w: 7.26,
    h: 0.4,
    fontSize: 16,
    bold: true,
    color: kliniOrange,
    align: "center",
  });

  const demoRows: any = [[
    { text: "FAIXA ETÁRIA", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    { text: "TITULAR M", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    { text: "TITULAR F", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    { text: "DEP. M", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    { text: "DEP. F", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    { text: "TOTAL M", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    { text: "TOTAL F", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    { text: "TOTAL", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    { text: "%", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
  ]];

  data.demographics.forEach((row, idx) => {
    const bgColor = idx % 2 === 0 ? white : kliniTealLight;
    demoRows.push([
      { text: row.ageRange, options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(row.titularM || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(row.titularF || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(row.dependentM || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(row.dependentF || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(row.totalM || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(row.totalF || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(row.total || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(row.percentage || "0%"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
    ]);
  });

  slide2.addTable(demoRows, {
    x: 0.3,
    y: 1,
    w: 7.66,
    border: { pt: 0.5, color: "CCCCCC" },
  });

  // ==================== SLIDE 3: SEM COPARTICIPAÇÃO ====================
  const slide3 = pptx.addSlide();
  slide3.background = { color: white };

  slide3.addText("PLANOS SEM COPARTICIPAÇÃO", {
    x: 0.5,
    y: 0.3,
    w: 7.26,
    h: 0.4,
    fontSize: 16,
    bold: true,
    color: kliniOrange,
    align: "center",
  });

  const noCopayRows: any = [[
    { text: "PLANO", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    { text: "CÓDIGO ANS", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    { text: "PER CAPITA", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    { text: "FATURA", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
  ]];

  data.plansWithoutCopay.forEach((plan, idx) => {
    const bgColor = idx % 2 === 0 ? white : kliniTealLight;
    noCopayRows.push([
      { text: plan.name, options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(plan.ansCode), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: formatCurrency(plan.perCapita), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: formatCurrency(plan.estimatedInvoice), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
    ]);
  });

  slide3.addTable(noCopayRows, {
    x: 0.5,
    y: 1,
    w: 7.26,
    border: { pt: 0.5, color: "CCCCCC" },
  });

  if (data.ageBasedPricingNoCopay && data.ageBasedPricingNoCopay.length > 0) {
    const planColumns = Object.keys(data.ageBasedPricingNoCopay[0]).filter((key) => key !== "ageRange");
    const ageRowsNoCopay: any = [[
      { text: "FAIXA ETÁRIA", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 7 } },
      ...planColumns.map((planName) => ({
        text: planName.substring(0, 12),
        options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 7 },
      })),
    ]];

    data.ageBasedPricingNoCopay.forEach((row, idx) => {
      const bgColor = idx % 2 === 0 ? white : kliniTealLight;
      ageRowsNoCopay.push([
        { text: row.ageRange, options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 7 } },
        ...planColumns.map((col) => ({
          text: row[col] ? formatCurrency(row[col]) : "-",
          options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 7 },
        })),
      ]);
    });

    const colWidth = (7.26 - 0.8) / planColumns.length;
    slide3.addTable(ageRowsNoCopay, {
      x: 0.5,
      y: 3,
      w: 7.26,
      border: { pt: 0.5, color: "CCCCCC" },
      colW: [0.8, ...Array(planColumns.length).fill(colWidth)],
    });
  }

  // ==================== SLIDE 4: COM COPARTICIPAÇÃO ====================
  const slide4 = pptx.addSlide();
  slide4.background = { color: white };

  slide4.addText("PLANOS COM COPARTICIPAÇÃO", {
    x: 0.5,
    y: 0.3,
    w: 7.26,
    h: 0.4,
    fontSize: 16,
    bold: true,
    color: kliniOrange,
    align: "center",
  });

  const copayRows: any = [[
    { text: "PLANO", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    { text: "CÓDIGO ANS", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    { text: "PER CAPITA", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    { text: "FATURA", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
  ]];

  data.plansWithCopay.forEach((plan, idx) => {
    const bgColor = idx % 2 === 0 ? white : kliniTealLight;
    copayRows.push([
      { text: plan.name, options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(plan.ansCode), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: formatCurrency(plan.perCapita), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: formatCurrency(plan.estimatedInvoice), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
    ]);
  });

  slide4.addTable(copayRows, {
    x: 0.5,
    y: 1,
    w: 7.26,
    border: { pt: 0.5, color: "CCCCCC" },
  });

  if (data.ageBasedPricingCopay && data.ageBasedPricingCopay.length > 0) {
    const planColumns = Object.keys(data.ageBasedPricingCopay[0]).filter((key) => key !== "ageRange");
    const ageRowsCopay: any = [[
      { text: "FAIXA ETÁRIA", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 7 } },
      ...planColumns.map((planName) => ({
        text: planName.substring(0, 12),
        options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 7 },
      })),
    ]];

    data.ageBasedPricingCopay.forEach((row, idx) => {
      const bgColor = idx % 2 === 0 ? white : kliniTealLight;
      ageRowsCopay.push([
        { text: row.ageRange, options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 7 } },
        ...planColumns.map((col) => ({
          text: row[col] ? formatCurrency(row[col]) : "-",
          options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 7 },
        })),
      ]);
    });

    const colWidth = (7.26 - 0.8) / planColumns.length;
    slide4.addTable(ageRowsCopay, {
      x: 0.5,
      y: 3,
      w: 7.26,
      border: { pt: 0.5, color: "CCCCCC" },
      colW: [0.8, ...Array(planColumns.length).fill(colWidth)],
    });
  }

  // ==================== SLIDES 5-7: PRODUTOS G ====================
  if (includeProductosG && data.demographicsG && data.demographicsG.length > 0) {
    // SLIDE 5
    const slide5 = pptx.addSlide();
    slide5.background = { color: white };

    slide5.addText("DEMOGRAFIA - PRODUTOS G", {
      x: 0.5,
      y: 0.3,
      w: 7.26,
      h: 0.4,
      fontSize: 16,
      bold: true,
      color: kliniOrange,
      align: "center",
    });

    const demGRows: any = [[
      { text: "FAIXA ETÁRIA", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
      { text: "TITULAR M", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
      { text: "TITULAR F", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
      { text: "DEP. M", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
      { text: "DEP. F", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
      { text: "TOTAL M", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
      { text: "TOTAL F", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
      { text: "TOTAL", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
      { text: "%", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    ]];

    data.demographicsG.forEach((row, idx) => {
      const bgColor = idx % 2 === 0 ? white : kliniTealLight;
      demGRows.push([
        { text: row.ageRange, options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        { text: String(row.titularM || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        { text: String(row.titularF || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        { text: String(row.dependentM || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        { text: String(row.dependentF || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        { text: String(row.totalM || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        { text: String(row.totalF || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        { text: String(row.total || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        { text: String(row.percentage || "0%"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      ]);
    });

    slide5.addTable(demGRows, {
      x: 0.3,
      y: 1,
      w: 7.66,
      border: { pt: 0.5, color: "CCCCCC" },
    });

    // SLIDE 6
    const slide6 = pptx.addSlide();
    slide6.background = { color: white };

    slide6.addText("PLANOS SEM COPAY - PRODUTOS G", {
      x: 0.5,
      y: 0.3,
      w: 7.26,
      h: 0.4,
      fontSize: 16,
      bold: true,
      color: kliniOrange,
      align: "center",
    });

    const noCopayGRows: any = [[
      { text: "PLANO", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
      { text: "CÓDIGO ANS", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
      { text: "PER CAPITA", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
      { text: "FATURA", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    ]];

    if (data.plansWithoutCopayG) {
      data.plansWithoutCopayG.forEach((plan, idx) => {
        const bgColor = idx % 2 === 0 ? white : kliniTealLight;
        noCopayGRows.push([
          { text: plan.name, options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
          { text: String(plan.ansCode), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
          { text: formatCurrency(plan.perCapita), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
          { text: formatCurrency(plan.estimatedInvoice), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        ]);
      });
    }

    slide6.addTable(noCopayGRows, {
      x: 0.5,
      y: 1,
      w: 7.26,
      border: { pt: 0.5, color: "CCCCCC" },
    });

    // SLIDE 7
    const slide7 = pptx.addSlide();
    slide7.background = { color: white };

    slide7.addText("PLANOS COM COPAY - PRODUTOS G", {
      x: 0.5,
      y: 0.3,
      w: 7.26,
      h: 0.4,
      fontSize: 16,
      bold: true,
      color: kliniOrange,
      align: "center",
    });

    const copayGRows: any = [[
      { text: "PLANO", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
      { text: "CÓDIGO ANS", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
      { text: "PER CAPITA", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
      { text: "FATURA", options: { fill: { color: kliniTeal }, color: white, bold: true, fontSize: 9 } },
    ]];

    if (data.plansWithCopayG) {
      data.plansWithCopayG.forEach((plan, idx) => {
        const bgColor = idx % 2 === 0 ? white : kliniTealLight;
        copayGRows.push([
          { text: plan.name, options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
          { text: String(plan.ansCode), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
          { text: formatCurrency(plan.perCapita), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
          { text: formatCurrency(plan.estimatedInvoice), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        ]);
      });
    }

    slide7.addTable(copayGRows, {
      x: 0.5,
      y: 1,
      w: 7.26,
      border: { pt: 0.5, color: "CCCCCC" },
    });
  }

  await pptx.writeFile({ fileName: `Proposta_${data.companyName || "Klini_Saude"}.pptx` });
};
