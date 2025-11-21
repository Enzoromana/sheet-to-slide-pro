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

const formatPercentage = (value: any): string => {
  if (typeof value === 'string' && value.includes('%')) {
    return value; // Already formatted
  }
  const num = typeof value === 'number' ? value : parseFloat(value);
  if (isNaN(num)) return '0%';
  return `${Math.round(num * 100)}%`;
};

export const exportToPPTX = async (data: ExportData, coverImage?: string | null) => {
  const pptx = new pptxgen();

  // Set portrait layout: 20.99cm x 29.704cm (8.26" x 11.69")
  pptx.defineLayout({ name: "PORTRAIT_A4", width: 8.26, height: 11.69 });
  pptx.layout = "PORTRAIT_A4";

  // Define cores Klini
  const kliniTeal = "1D7874";
  const kliniOrange = "F7931E";
  const lightTeal = "B8D4D3";

  // Slide 1: Capa com imagem de fundo ou design padrão
  const slide1 = pptx.addSlide();
  
  if (coverImage) {
    // Se o usuário fez upload de uma imagem, usar ela como capa completa
    slide1.addImage({
      data: coverImage,
      x: 0,
      y: 0,
      w: 8.26,
      h: 11.69,
      sizing: { type: "cover" }
    });
  } else {
    // Caso contrário, usar o design padrão
    // Background gradient para simular a capa Klini
    slide1.background = { color: "1D7874" };
    
    // Adicionar elementos da capa manualmente para replicar o design
    // Círculo decorativo (simulando o Rio de Janeiro ao fundo)
    slide1.addShape(pptx.ShapeType.ellipse, {
      x: 1.5,
      y: 3.0,
      w: 5.0,
      h: 5.0,
      fill: { color: "164E4B", transparency: 30 },
      line: { type: "none" }
    });

    // Logo "klini saúde" no topo
    slide1.addText([
      { text: "klini", options: { fontSize: 72, bold: true, color: "FFFFFF" } }
    ], {
      x: 1.8,
      y: 1.8,
      w: 4.66,
      h: 1.5,
      align: "center"
    });
    
    slide1.addText([
      { text: "saúde", options: { fontSize: 48, bold: false, color: "FFFFFF" } }
    ], {
      x: 2.3,
      y: 2.8,
      w: 3.66,
      h: 1.0,
      align: "center"
    });

    // Caixa branca "PROPOSTA COMERCIAL"
    slide1.addShape(pptx.ShapeType.rect, {
      x: 0.5,
      y: 6.8,
      w: 7.26,
      h: 1.0,
      fill: { color: "FFFFFF" },
      line: { type: "none" }
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
      valign: "middle"
    });

    // Caixa laranja com data
    const currentDate = new Date().toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
    slide1.addShape(pptx.ShapeType.rect, {
      x: 4.7,
      y: 8.0,
      w: 2.8,
      h: 0.6,
      fill: { color: kliniOrange },
      line: { type: "none" }
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
      italic: true
    });

    // Versão no rodapé esquerdo
    slide1.addText("V2.00/070251.3.4", {
      x: 0.3,
      y: 11.35,
      w: 2.0,
      h: 0.25,
      fontSize: 7,
      color: "FFFFFF",
      align: "left",
    });
    
    // ANS no rodapé direito
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

  // Slide 2: Informações da Empresa + Demografia
  const slide2 = pptx.addSlide();
  slide2.background = { color: "FFFFFF" };

  // Título
  slide2.addText([
    { text: "Tabela de ", options: { color: kliniOrange, fontSize: 28, bold: true } },
    { text: "Preços", options: { color: kliniOrange, fontSize: 28, bold: true } },
  ], {
    x: 0.5,
    y: 0.5,
    w: 7.26,
    h: 0.5,
    align: "center",
  });

  slide2.addText("DEMOGRAFIA", {
    x: 0.5,
    y: 1.1,
    w: 7.26,
    h: 0.3,
    fontSize: 12,
    color: "666666",
    align: "center",
  });

  // Informações da empresa
  slide2.addText(`Razão Social: ${data.companyName}`, {
    x: 0.5,
    y: 1.6,
    w: 7.26,
    h: 0.3,
    fontSize: 10,
    color: "333333",
  });

  slide2.addText(`Concessionária: ${data.concessionaire}`, {
    x: 0.5,
    y: 1.9,
    w: 7.26,
    h: 0.3,
    fontSize: 10,
    color: "333333",
  });

  slide2.addText(`Corretor(a): ${data.broker}`, {
    x: 0.5,
    y: 2.2,
    w: 7.26,
    h: 0.3,
    fontSize: 10,
    color: "333333",
  });

  // Tabela de Demografia
  const demoRows = [
    [
      { text: "FAIXA ETÁRIA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, rowspan: 2 } },
      { text: "TITULAR", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, colspan: 2 } },
      { text: "DEPENDENTE", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, colspan: 2 } },
      { text: "AGREGADO", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, colspan: 2 } },
      { text: "TOTAL", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, colspan: 3 } },
      { text: "%", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal }, rowspan: 2 } },
    ],
    [
      { text: "M", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
      { text: "F", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
      { text: "M", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
      { text: "F", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
      { text: "M", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
      { text: "F", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
      { text: "M", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
      { text: "F", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
      { text: "TOTAL", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
    ],
  ];

  data.demographics.forEach((row, index) => {
    const isAlt = index % 2 === 1;
    demoRows.push([
      { text: row.ageRange, options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: String(row.titularM || '0'), options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: String(row.titularF || '0'), options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: String(row.dependentM || '0'), options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: String(row.dependentF || '0'), options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: String(row.agregadoM || '0'), options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: String(row.agregadoF || '0'), options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: String(row.totalM || '0'), options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: String(row.totalF || '0'), options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: String(row.total), options: { bold: true, fill: { color: kliniTeal }, color: "FFFFFF" } },
      { text: formatPercentage(row.percentage), options: { bold: true, fill: { color: kliniTeal }, color: "FFFFFF" } },
    ]);
  });

  slide2.addTable(demoRows, {
    x: 0.3,
    y: 2.7,
    w: 7.66,
    fontSize: 7,
    border: { pt: 0.5, color: "CCCCCC" },
    align: "center",
  });

  // Slide 3: Planos com Coparticipação
  const slide3 = pptx.addSlide();
  slide3.background = { color: "FFFFFF" };

  slide3.addText("Planos com Coparticipação", {
    x: 0.5,
    y: 0.5,
    w: 7.26,
    h: 0.5,
    fontSize: 24,
    bold: true,
    color: kliniOrange,
    align: "center",
  });

  const copayRows = [
    [
      { text: "PLANO", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
      { text: "Registro ANS", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
      { text: "Valor Per Capita", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
      { text: "Fatura Estimada", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
    ],
  ];

  const formatCurrency = (value: string | number) => {
    const numValue = typeof value === 'string' ? parseFloat(value.replace(/[^\d,.-]/g, '').replace(',', '.')) : value;
    return `R$ ${numValue.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  };

  data.plansWithCopay.forEach((plan, index) => {
    const isAlt = index % 2 === 1;
    copayRows.push([
      { text: plan.name, options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: String(plan.ansCode), options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: formatCurrency(plan.perCapita), options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: formatCurrency(plan.estimatedInvoice), options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
    ]);
  });

  slide3.addTable(copayRows, {
    x: 0.5,
    y: 1.3,
    w: 7.26,
    fontSize: 9,
    border: { pt: 0.5, color: "CCCCCC" },
    align: "center",
  });

  // Slide 4: Planos sem Coparticipação
  const slide4 = pptx.addSlide();
  slide4.background = { color: "FFFFFF" };

  slide4.addText("Planos sem Coparticipação", {
    x: 0.5,
    y: 0.5,
    w: 7.26,
    h: 0.5,
    fontSize: 24,
    bold: true,
    color: kliniOrange,
    align: "center",
  });

  const noCopayRows = [
    [
      { text: "PLANO", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
      { text: "Registro ANS", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
      { text: "Valor Per Capita", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
      { text: "Fatura Estimada", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
    ],
  ];

  data.plansWithoutCopay.forEach((plan, index) => {
    const isAlt = index % 2 === 1;
    noCopayRows.push([
      { text: plan.name, options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: String(plan.ansCode), options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: formatCurrency(plan.perCapita), options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
      { text: formatCurrency(plan.estimatedInvoice), options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
    ]);
  });

  slide4.addTable(noCopayRows, {
    x: 0.5,
    y: 1.3,
    w: 7.26,
    fontSize: 9,
    border: { pt: 0.5, color: "CCCCCC" },
    align: "center",
  });

  // Slides de Valores por Faixa Etária - Planos COM Coparticipação
  if (data.ageBasedPricingCopay && data.ageBasedPricingCopay.length > 0) {
    const planColumns = Object.keys(data.ageBasedPricingCopay[0]).filter(key => key !== 'ageRange');
    
    // Dividir planos em grupos de 3 para formato retrato
    const plansPerSlide = 3;
    for (let i = 0; i < planColumns.length; i += plansPerSlide) {
      const columnsSlice = planColumns.slice(i, i + plansPerSlide);
      
      const slideAge = pptx.addSlide();
      slideAge.background = { color: "FFFFFF" };

      slideAge.addText("Valores por Faixa Etária - COM Coparticipação", {
        x: 0.5,
        y: 0.5,
        w: 7.26,
        h: 0.5,
        fontSize: 20,
        bold: true,
        color: kliniOrange,
        align: "center",
      });

      const ageRows = [
        [
          { text: "FAIXA ETÁRIA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
          ...columnsSlice.map((planName) => ({
            text: planName.length > 30 ? planName.substring(0, 30) + "..." : planName,
            options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } }
          }))
        ],
      ];

      data.ageBasedPricingCopay.forEach((row, rowIndex) => {
        const isAlt = rowIndex % 2 === 1;
        ageRows.push([
          { text: row.ageRange, options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
          ...columnsSlice.map(col => ({
            text: row[col] ? formatCurrency(row[col]) : '-',
            options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } }
          }))
        ]);
      });

      slideAge.addTable(ageRows, {
        x: 0.3,
        y: 1.2,
        w: 7.66,
        fontSize: 7,
        border: { pt: 0.5, color: "CCCCCC" },
        align: "center",
        colW: [1.5, ...Array(columnsSlice.length).fill((7.66 - 1.5) / columnsSlice.length)]
      });

      slideAge.addText(
        "Esta proposta foi elaborada levando em consideração as informações fornecidas através do formulário de cotação enviado pela Corretora. No caso de implantação do contrato, qualquer incompatibilidade implicará na inviabilidade ou reanálise da proposta.",
        {
          x: 0.5,
          y: 10.8,
          w: 7.26,
          h: 0.6,
          fontSize: 7,
          color: "666666",
          align: "center",
          valign: "top",
        }
      );
      
      slideAge.addText("ANS - Nº 42.202-9", {
        x: 6.5,
        y: 0.2,
        w: 1.5,
        h: 0.3,
        fontSize: 9,
        color: "333333",
        align: "right",
      });
    }
  }

  // Slides de Valores por Faixa Etária - Planos SEM Coparticipação
  if (data.ageBasedPricingNoCopay && data.ageBasedPricingNoCopay.length > 0) {
    const planColumns = Object.keys(data.ageBasedPricingNoCopay[0]).filter(key => key !== 'ageRange');
    
    // Dividir planos em grupos de 3 para formato retrato
    const plansPerSlide = 3;
    for (let i = 0; i < planColumns.length; i += plansPerSlide) {
      const columnsSlice = planColumns.slice(i, i + plansPerSlide);
      
      const slideAge = pptx.addSlide();
      slideAge.background = { color: "FFFFFF" };

      slideAge.addText("Valores por Faixa Etária - SEM Coparticipação", {
        x: 0.5,
        y: 0.5,
        w: 7.26,
        h: 0.5,
        fontSize: 20,
        bold: true,
        color: kliniOrange,
        align: "center",
      });

      const ageRows = [
        [
          { text: "FAIXA ETÁRIA", options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } } },
          ...columnsSlice.map((planName) => ({
            text: planName.length > 30 ? planName.substring(0, 30) + "..." : planName,
            options: { bold: true, color: "FFFFFF", fill: { color: kliniTeal } }
          }))
        ],
      ];

      data.ageBasedPricingNoCopay.forEach((row, rowIndex) => {
        const isAlt = rowIndex % 2 === 1;
        ageRows.push([
          { text: row.ageRange, options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } } },
          ...columnsSlice.map(col => ({
            text: row[col] ? formatCurrency(row[col]) : '-',
            options: { bold: false, color: "333333", fill: { color: isAlt ? lightTeal : "FFFFFF" } }
          }))
        ]);
      });

      slideAge.addTable(ageRows, {
        x: 0.3,
        y: 1.2,
        w: 7.66,
        fontSize: 7,
        border: { pt: 0.5, color: "CCCCCC" },
        align: "center",
        colW: [1.5, ...Array(columnsSlice.length).fill((7.66 - 1.5) / columnsSlice.length)]
      });

      slideAge.addText(
        "Esta proposta foi elaborada levando em consideração as informações fornecidas através do formulário de cotação enviado pela Corretora. No caso de implantação do contrato, qualquer incompatibilidade implicará na inviabilidade ou reanálise da proposta.",
        {
          x: 0.5,
          y: 10.8,
          w: 7.26,
          h: 0.6,
          fontSize: 7,
          color: "666666",
          align: "center",
          valign: "top",
        }
      );
      
      slideAge.addText("ANS - Nº 42.202-9", {
        x: 6.5,
        y: 0.2,
        w: 1.5,
        h: 0.3,
        fontSize: 9,
        color: "333333",
        align: "right",
      });
    }
  }

  // Adicionar nota de rodapé em todos os slides de conteúdo
  [slide2, slide3, slide4].forEach(slide => {
    slide.addText(
      "Esta proposta foi elaborada levando em consideração as informações fornecidas através do formulário de cotação enviado pela Corretora. No caso de implantação do contrato, qualquer incompatibilidade implicará na inviabilidade ou reanálise da proposta.",
      {
        x: 0.5,
        y: 10.8,
        w: 7.26,
        h: 0.6,
        fontSize: 7,
        color: "666666",
        align: "center",
        valign: "top",
      }
    );
    
    slide.addText("ANS - Nº 42.202-9", {
      x: 6.5,
      y: 0.2,
      w: 1.5,
      h: 0.3,
      fontSize: 9,
      color: "333333",
      align: "right",
    });
  });

  // Salvar arquivo
  await pptx.writeFile({ fileName: `Proposta_${data.companyName || 'Klini_Saude'}.pptx` });
};
