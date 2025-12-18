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
    color: "FFFFFF",
    align: "center",
  });

  slide1.addText("saúde", {
    x: 2,
    y: 5,
    w: 4,
    h: 1,
    fontSize: 48,
    bold: false,
    color: "FFFFFF",
    align: "center",
  });

  // ==================== SLIDE 2: DEMOGRAFIA ====================
  const slide2 = pptx.addSlide();
  slide2.background = { color: "FFFFFF" };

  slide2.addText("DEMOGRAFIA - TESTE DE CORES", {
    x: 0.5,
    y: 0.3,
    w: 7.26,
    h: 0.4,
    fontSize: 16,
    bold: true,
    color: kliniOrange,
    align: "center",
  });

  // Criar tabela com ESTRUCTURA SIMPLES usando rowH e tableStyleId
  const tableData: any = [];

  // HEADER
  tableData.push([
    {
      text: "FAIXA ETÁRIA",
      options: {
        fill: kliniTeal,
        color: "FFFFFF",
        bold: true,
        fontSize: 9,
        align: "center",
      },
    },
    {
      text: "TITULAR M",
      options: {
        fill: kliniTeal,
        color: "FFFFFF",
        bold: true,
        fontSize: 9,
        align: "center",
      },
    },
    {
      text: "TITULAR F",
      options: {
        fill: kliniTeal,
        color: "FFFFFF",
        bold: true,
        fontSize: 9,
        align: "center",
      },
    },
    {
      text: "TOTAL",
      options: {
        fill: kliniTeal,
        color: "FFFFFF",
        bold: true,
        fontSize: 9,
        align: "center",
      },
    },
  ]);

  // DATA ROWS
  data.demographics.forEach((row, idx) => {
    const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
    tableData.push([
      {
        text: row.ageRange,
        options: {
          fill: bgColor,
          color: "333333",
          bold: false,
          fontSize: 8,
          align: "center",
        },
      },
      {
        text: String(row.titularM || "0"),
        options: {
          fill: bgColor,
          color: "333333",
          bold: false,
          fontSize: 8,
          align: "center",
        },
      },
      {
        text: String(row.titularF || "0"),
        options: {
          fill: bgColor,
          color: "333333",
          bold: false,
          fontSize: 8,
          align: "center",
        },
      },
      {
        text: String(row.total || "0"),
        options: {
          fill: bgColor,
          color: "333333",
          bold: false,
          fontSize: 8,
          align: "center",
        },
      },
    ]);
  });

  slide2.addTable(tableData, {
    x: 1,
    y: 1,
    w: 6.26,
    align: "center",
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
    {
      text: "PLANO",
      options: {
        fill: kliniTeal,
        color: "FFFFFF",
        bold: true,
        fontSize: 9,
        align: "center",
      },
    },
    {
      text: "CÓDIGO ANS",
      options: {
        fill: kliniTeal,
        color: "FFFFFF",
        bold: true,
        fontSize: 9,
        align: "center",
      },
    },
    {
      text: "PER CAPITA",
      options: {
        fill: kliniTeal,
        color: "FFFFFF",
        bold: true,
        fontSize: 9,
        align: "center",
      },
    },
    {
      text: "TOTAL",
      options: {
        fill: kliniTeal,
        color: "FFFFFF",
        bold: true,
        fontSize: 9,
        align: "center",
      },
    },
  ]);

  data.plansWithoutCopay.forEach((plan, idx) => {
    const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
    noCopayData.push([
      {
        text: plan.name,
        options: {
          fill: bgColor,
          color: "333333",
          bold: false,
          fontSize: 8,
          align: "center",
        },
      },
      {
        text: String(plan.ansCode),
        options: {
          fill: bgColor,
          color: "333333",
          bold: false,
          fontSize: 8,
          align: "center",
        },
      },
      {
        text: formatCurrency(plan.perCapita),
        options: {
          fill: bgColor,
          color: "333333",
          bold: false,
          fontSize: 8,
          align: "center",
        },
      },
      {
        text: formatCurrency(plan.estimatedInvoice),
        options: {
          fill: bgColor,
          color: "333333",
          bold: false,
          fontSize: 8,
          align: "center",
        },
      },
    ]);
  });

  slide3.addTable(noCopayData, {
    x: 1,
    y: 1,
    w: 6.26,
    align: "center",
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
    {
      text: "PLANO",
      options: {
        fill: kliniTeal,
        color: "FFFFFF",
        bold: true,
        fontSize: 9,
        align: "center",
      },
    },
    {
      text: "CÓDIGO ANS",
      options: {
        fill: kliniTeal,
        color: "FFFFFF",
        bold: true,
        fontSize: 9,
        align: "center",
      },
    },
    {
      text: "PER CAPITA",
      options: {
        fill: kliniTeal,
        color: "FFFFFF",
        bold: true,
        fontSize: 9,
        align: "center",
      },
    },
    {
      text: "TOTAL",
      options: {
        fill: kliniTeal,
        color: "FFFFFF",
        bold: true,
        fontSize: 9,
        align: "center",
      },
    },
  ]);

  data.plansWithCopay.forEach((plan, idx) => {
    const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
    copayData.push([
      {
        text: plan.name,
        options: {
          fill: bgColor,
          color: "333333",
          bold: false,
          fontSize: 8,
          align: "center",
        },
      },
      {
        text: String(plan.ansCode),
        options: {
          fill: bgColor,
          color: "333333",
          bold: false,
          fontSize: 8,
          align: "center",
        },
      },
      {
        text: formatCurrency(plan.perCapita),
        options: {
          fill: bgColor,
          color: "333333",
          bold: false,
          fontSize: 8,
          align: "center",
        },
      },
      {
        text: formatCurrency(plan.estimatedInvoice),
        options: {
          fill: bgColor,
          color: "333333",
          bold: false,
          fontSize: 8,
          align: "center",
        },
      },
    ]);
  });

  slide4.addTable(copayData, {
    x: 1,
    y: 1,
    w: 6.26,
    align: "center",
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
      {
        text: "FAIXA ETÁRIA",
        options: {
          fill: kliniTeal,
          color: "FFFFFF",
          bold: true,
          fontSize: 9,
          align: "center",
        },
      },
      {
        text: "TITULAR M",
        options: {
          fill: kliniTeal,
          color: "FFFFFF",
          bold: true,
          fontSize: 9,
          align: "center",
        },
      },
      {
        text: "TITULAR F",
        options: {
          fill: kliniTeal,
          color: "FFFFFF",
          bold: true,
          fontSize: 9,
          align: "center",
        },
      },
      {
        text: "TOTAL",
        options: {
          fill: kliniTeal,
          color: "FFFFFF",
          bold: true,
          fontSize: 9,
          align: "center",
        },
      },
    ]);

    data.demographicsG.forEach((row, idx) => {
      const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
      demGData.push([
        {
          text: row.ageRange,
          options: {
            fill: bgColor,
            color: "333333",
            bold: false,
            fontSize: 8,
            align: "center",
          },
        },
        {
          text: String(row.titularM || "0"),
          options: {
            fill: bgColor,
            color: "333333",
            bold: false,
            fontSize: 8,
            align: "center",
          },
        },
        {
          text: String(row.titularF || "0"),
          options: {
            fill: bgColor,
            color: "333333",
            bold: false,
            fontSize: 8,
            align: "center",
          },
        },
        {
          text: String(row.total || "0"),
          options: {
            fill: bgColor,
            color: "333333",
            bold: false,
            fontSize: 8,
            align: "center",
          },
        },
      ]);
    });

    slide5.addTable(demGData, {
      x: 1,
      y: 1,
      w: 6.26,
      align: "center",
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
      {
        text: "PLANO",
        options: {
          fill: kliniTeal,
          color: "FFFFFF",
          bold: true,
          fontSize: 9,
          align: "center",
        },
      },
      {
        text: "CÓDIGO ANS",
        options: {
          fill: kliniTeal,
          color: "FFFFFF",
          bold: true,
          fontSize: 9,
          align: "center",
        },
      },
      {
        text: "PER CAPITA",
        options: {
          fill: kliniTeal,
          color: "FFFFFF",
          bold: true,
          fontSize: 9,
          align: "center",
        },
      },
      {
        text: "TOTAL",
        options: {
          fill: kliniTeal,
          color: "FFFFFF",
          bold: true,
          fontSize: 9,
          align: "center",
        },
      },
    ]);

    if (data.plansWithoutCopayG && data.plansWithoutCopayG.length > 0) {
      data.plansWithoutCopayG.forEach((plan, idx) => {
        const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
        noCopayGData.push([
          {
            text: plan.name,
            options: {
              fill: bgColor,
              color: "333333",
              bold: false,
              fontSize: 8,
              align: "center",
            },
          },
          {
            text: String(plan.ansCode),
            options: {
              fill: bgColor,
              color: "333333",
              bold: false,
              fontSize: 8,
              align: "center",
            },
          },
          {
            text: formatCurrency(plan.perCapita),
            options: {
              fill: bgColor,
              color: "333333",
              bold: false,
              fontSize: 8,
              align: "center",
            },
          },
          {
            text: formatCurrency(plan.estimatedInvoice),
            options: {
              fill: bgColor,
              color: "333333",
              bold: false,
              fontSize: 8,
              align: "center",
            },
          },
        ]);
      });
    }

    slide6.addTable(noCopayGData, {
      x: 1,
      y: 1,
      w: 6.26,
      align: "center",
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
      {
        text: "PLANO",
        options: {
          fill: kliniTeal,
          color: "FFFFFF",
          bold: true,
          fontSize: 9,
          align: "center",
        },
      },
      {
        text: "CÓDIGO ANS",
        options: {
          fill: kliniTeal,
          color: "FFFFFF",
          bold: true,
          fontSize: 9,
          align: "center",
        },
      },
      {
        text: "PER CAPITA",
        options: {
          fill: kliniTeal,
          color: "FFFFFF",
          bold: true,
          fontSize: 9,
          align: "center",
        },
      },
      {
        text: "TOTAL",
        options: {
          fill: kliniTeal,
          color: "FFFFFF",
          bold: true,
          fontSize: 9,
          align: "center",
        },
      },
    ]);

    if (data.plansWithCopayG && data.plansWithCopayG.length > 0) {
      data.plansWithCopayG.forEach((plan, idx) => {
        const bgColor = idx % 2 === 0 ? "FFFFFF" : lightGray;
        copayGData.push([
          {
            text: plan.name,
            options: {
              fill: bgColor,
              color: "333333",
              bold: false,
              fontSize: 8,
              align: "center",
            },
          },
          {
            text: String(plan.ansCode),
            options: {
              fill: bgColor,
              color: "333333",
              bold: false,
              fontSize: 8,
              align: "center",
            },
          },
          {
            text: formatCurrency(plan.perCapita),
            options: {
              fill: bgColor,
              color: "333333",
              bold: false,
              fontSize: 8,
              align: "center",
            },
          },
          {
            text: formatCurrency(plan.estimatedInvoice),
            options: {
              fill: bgColor,
              color: "333333",
              bold: false,
              fontSize: 8,
              align: "center",
            },
          },
        ]);
      });
    }

    slide7.addTable(copayGData, {
      x: 1,
      y: 1,
      w: 6.26,
      align: "center",
      border: { pt: 1, color: "CCCCCC" },
    });
  }

  await pptx.writeFile({ fileName: `Proposta_${data.companyName || "Klini_Saude"}.pptx` });
};
