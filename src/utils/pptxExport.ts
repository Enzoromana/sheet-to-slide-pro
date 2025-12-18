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

  // USAR DIRETAMENTE COMO STRING HEX
  const kliniTeal = "1D7874";
  const kliniOrange = "F7931E";
  const lightGray = "F5F5F5";

  // ==================== SLIDE 1: CAPA ====================
  const slide1 = pptx.addSlide();
  slide1.background = { color: kliniTeal };
  
  slide1.addText("klini saúde", {
    x: 2,
    y: 5,
    w: 4,
    h: 1.5,
    fontSize: 60,
    bold: true,
    color: "FFFFFF",
    align: "center",
  });

  // ==================== SLIDE 2: DEMOGRAFIA ====================
  const slide2 = pptx.addSlide();
  slide2.background = { color: "FFFFFF" };

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

  const tableData: any = [];

  // HEADER
  tableData.push([
    { text: "FAIXA ETÁRIA", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
    { text: "TITULAR M", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
    { text: "TITULAR F", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
    { text: "TOTAL", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
  ]);

  // DATA
  data.demographics.forEach((row, idx) => {
    const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
    tableData.push([
      { text: row.ageRange, options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(row.titularM || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(row.titularF || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(row.total || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
    ]);
  });

  slide2.addTable(tableData, {
    x: 0.5,
    y: 1,
    w: 7.26,
    border: { pt: 1, color: "CCCCCC" },
  });

  // ==================== SLIDE 3: PLANOS SEM COPAY ====================
  const slide3 = pptx.addSlide();
  slide3.background = { color: "FFFFFF" };

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

  const noCopayData: any = [];
  noCopayData.push([
    { text: "PLANO", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
    { text: "CÓDIGO ANS", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
    { text: "PER CAPITA", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
    { text: "TOTAL", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
  ]);

  data.plansWithoutCopay.forEach((plan, idx) => {
    const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
    noCopayData.push([
      { text: plan.name, options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(plan.ansCode), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: formatCurrency(plan.perCapita), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: formatCurrency(plan.estimatedInvoice), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
    ]);
  });

  slide3.addTable(noCopayData, {
    x: 0.5,
    y: 1,
    w: 7.26,
    border: { pt: 1, color: "CCCCCC" },
  });

  // ==================== SLIDE 4: PLANOS COM COPAY ====================
  const slide4 = pptx.addSlide();
  slide4.background = { color: "FFFFFF" };

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

  const copayData: any = [];
  copayData.push([
    { text: "PLANO", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
    { text: "CÓDIGO ANS", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
    { text: "PER CAPITA", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
    { text: "TOTAL", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
  ]);

  data.plansWithCopay.forEach((plan, idx) => {
    const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
    copayData.push([
      { text: plan.name, options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: String(plan.ansCode), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: formatCurrency(plan.perCapita), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      { text: formatCurrency(plan.estimatedInvoice), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
    ]);
  });

  slide4.addTable(copayData, {
    x: 0.5,
    y: 1,
    w: 7.26,
    border: { pt: 1, color: "CCCCCC" },
  });

  // ==================== SLIDES 5-7: PRODUTOS G ====================
  if (includeProductosG && data.demographicsG && data.demographicsG.length > 0) {
    // SLIDE 5
    const slide5 = pptx.addSlide();
    slide5.background = { color: "FFFFFF" };

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

    const demGData: any = [];
    demGData.push([
      { text: "FAIXA ETÁRIA", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
      { text: "TITULAR M", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
      { text: "TITULAR F", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
      { text: "TOTAL", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
    ]);

    data.demographicsG.forEach((row, idx) => {
      const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
      demGData.push([
        { text: row.ageRange, options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        { text: String(row.titularM || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        { text: String(row.titularF || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        { text: String(row.total || "0"), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
      ]);
    });

    slide5.addTable(demGData, {
      x: 0.5,
      y: 1,
      w: 7.26,
      border: { pt: 1, color: "CCCCCC" },
    });

    // SLIDE 6
    const slide6 = pptx.addSlide();
    slide6.background = { color: "FFFFFF" };

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

    const noCopayGData: any = [];
    noCopayGData.push([
      { text: "PLANO", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
      { text: "CÓDIGO ANS", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
      { text: "PER CAPITA", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
      { text: "TOTAL", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
    ]);

    if (data.plansWithoutCopayG && data.plansWithoutCopayG.length > 0) {
      data.plansWithoutCopayG.forEach((plan, idx) => {
        const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
        noCopayGData.push([
          { text: plan.name, options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
          { text: String(plan.ansCode), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
          { text: formatCurrency(plan.perCapita), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
          { text: formatCurrency(plan.estimatedInvoice), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        ]);
      });
    }

    slide6.addTable(noCopayGData, {
      x: 0.5,
      y: 1,
      w: 7.26,
      border: { pt: 1, color: "CCCCCC" },
    });

    // SLIDE 7
    const slide7 = pptx.addSlide();
    slide7.background = { color: "FFFFFF" };

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

    const copayGData: any = [];
    copayGData.push([
      { text: "PLANO", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
      { text: "CÓDIGO ANS", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
      { text: "PER CAPITA", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
      { text: "TOTAL", options: { fill: { color: kliniTeal }, color: "FFFFFF", bold: true, fontSize: 9 } },
    ]);

    if (data.plansWithCopayG && data.plansWithCopayG.length > 0) {
      data.plansWithCopayG.forEach((plan, idx) => {
        const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
        copayGData.push([
          { text: plan.name, options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
          { text: String(plan.ansCode), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
          { text: formatCurrency(plan.perCapita), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
          { text: formatCurrency(plan.estimatedInvoice), options: { fill: { color: bgColor }, color: "333333", bold: false, fontSize: 8 } },
        ]);
      });
    }

    slide7.addTable(copayGData, {
      x: 0.5,
      y: 1,
      w: 7.26,
      border: { pt: 1, color: "CCCCCC" },
    });
  }

  await pptx.writeFile({ fileName: `Proposta_${data.companyName || "Klini_Saude"}.pptx` });
};
