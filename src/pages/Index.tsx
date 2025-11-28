import { useState } from "react";
import { Upload, Download, Presentation } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { useToast } from "@/hooks/use-toast";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";
import { PricingTable } from "@/components/PricingTable";
import { DemographicsTable } from "@/components/DemographicsTable";
import { CompanyHeader } from "@/components/CompanyHeader";
import { AgeBasedPricingTable } from "@/components/AgeBasedPricingTable";
import { CoverImageUpload } from "@/components/CoverImageUpload";
import { exportToPPTX } from "@/utils/pptxExport";
import { DEFAULT_COVER_IMAGE } from "@/utils/defaultCoverImage";
import { KLINI_LOGO } from "@/assets/kliniLogo";

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
}

const Index = () => {
  const [parsedData, setParsedData] = useState<ParsedData | null>(null);
  const [coverImage, setCoverImage] = useState<string | null>(DEFAULT_COVER_IMAGE);
  const { toast } = useToast();
  const [validationError, setValidationError] = useState<{
  errors: string[];
  warnings: string[];
} | null>(null);


  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const parseCurrency = (value: any): number => {
      if (typeof value === 'number') return value;
      if (typeof value === 'string') {
        const cleaned = value.replace(/[R$\s.]/g, '').replace(',', '.');
        const parsed = parseFloat(cleaned);
        return isNaN(parsed) ? 0 : parsed;
      }
      return 0;
    };

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

      // ✅ VALIDAR ESTRUTURA DO ARQUIVO
      const validation = validateExcelStructure(jsonData, workbook.SheetNames[0]);

      if (!validation.isValid || validation.warnings.length > 0) {
        setValidationError({
          errors: validation.errors,
          warnings: validation.warnings,
        });
        setShowValidationDialog(true);
        setPendingFile(data);
        setPendingWorkbook(workbook);
        return; // Parar aqui
      }
      // Parse company info - data is in column B (index 1), rows 2-4 (indices 1-3)
      const companyName = (jsonData[1] as any)?.[1] || "";
      const concessionaire = (jsonData[2] as any)?.[1] || "";
      const broker = (jsonData[3] as any)?.[1] || "";

      // Parse demographics (rows 8-17 in Excel = indices 7-16)
      // Column layout: B=age range, C-D=titular M/F, E-F=dependente M/F, I-J=total M/F, K=total, L=percentage
      const allDemographics: any[] = [];
      for (let i = 7; i <= 16; i++) {
        const row = jsonData[i] as any[];
        if (!row || !row[1]) continue;
        const ageRange = String(row[1]).trim();
        if (ageRange === "TOTAL" || ageRange === "IDADE MÉDIA:" || ageRange === "") continue;
        allDemographics.push({
          ageRange,
          titularM: row[2] || 0,
          titularF: row[3] || 0,
          dependentM: row[4] || 0,
          dependentF: row[5] || 0,
          agregadoM: 0,
          agregadoF: 0,
          totalM: row[8] || 0,
          totalF: row[9] || 0,
          total: row[10] || 0,
          percentage: row[11] || "0%",
        });
      }

      // Parse plans WITH copay (rows 27-34 in Excel = indices 26-33)
      // Column layout: B=name, E=ANS code, F=per capita, G=estimated invoice
      const allPlansWithCopay: any[] = [];
      for (let i = 26; i <= 33; i++) {
        const row = jsonData[i] as any[];
        if (!row || !row[1]) continue;
        const planName = String(row[1]).trim();
        if (!planName.startsWith("KLINI")) continue;
        allPlansWithCopay.push({
          name: planName,
          ansCode: String(row[4] || '').trim(),
          perCapita: parseCurrency(row[5]) || 0,
          estimatedInvoice: parseCurrency(row[6]) || 0,
        });
      }

      // Parse age-based pricing WITH copay (header in row 40 = index 39, data rows 43-52 = indices 42-51)
      // Header starts at column 2 (index 1)
      const copayAgeHeader = jsonData[40] as any[];
      const copayPlanNames = copayAgeHeader.slice(2).filter((name: string) => name && String(name).trim() !== "");
      const allAgeBasedPricingCopay: any[] = [];
        for (let i = 41; i <= 50; i++) {
        const row = jsonData[i] as any[];
        if (!row || !row[1]) continue;
        const ageRange = String(row[1]).trim();
        if (!ageRange.match(/\d+\s*-\s*\d+|59\+/)) continue;
        const pricing: any = { ageRange };
        copayPlanNames.forEach((planName: string, idx: number) => {
          pricing[planName] = parseCurrency(row[idx + 2]) || 0;
        });
        allAgeBasedPricingCopay.push(pricing);
      }

      // Parse plans WITHOUT copay (rows 61-68 in Excel = indices 60-67)
      // Column layout: B=name, E=ANS code, F=per capita, G=estimated invoice
      const allPlansWithoutCopay: any[] = [];
      for (let i = 60; i <= 67; i++) {
        const row = jsonData[i] as any[];
        if (!row || !row[1]) continue;
        const planName = String(row[1]).trim();
        if (!planName.startsWith("KLINI")) continue;
        allPlansWithoutCopay.push({
          name: planName,
          ansCode: String(row[4] || '').trim(),
          perCapita: parseCurrency(row[5]) || 0,
          estimatedInvoice: parseCurrency(row[6]) || 0,
        });
      }

      // Parse age-based pricing WITHOUT copay (header in row 75 = index 74, data rows 78-87 = indices 77-86)
      // Header starts at column 2 (index 1)
      const noCopayAgeHeader = jsonData[75] as any[];
      const noCopayPlanNames = noCopayAgeHeader.slice(2).filter((name: string) => name && String(name).trim() !== "");
      const allAgeBasedPricingNoCopay: any[] = [];
      for (let i = 76; i <= 85; i++) {
        const row = jsonData[i] as any[];
        if (!row || !row[1]) continue;
        const ageRange = String(row[1]).trim();
        if (!ageRange.match(/\d+\s*-\s*\d+|59\+/)) continue;
        const pricing: any = { ageRange };
        noCopayPlanNames.forEach((planName: string, idx: number) => {
          pricing[planName] = parseCurrency(row[idx + 2]) || 0;
        });
        allAgeBasedPricingNoCopay.push(pricing);
      }

      const emissionDate = new Date().toLocaleDateString('pt-BR');
      const validityDate = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toLocaleDateString('pt-BR');

      setParsedData({
        companyName,
        concessionaire,
        broker,
        emissionDate,
        validityDate,
        demographics: allDemographics,
        plansWithCopay: allPlansWithCopay,
        plansWithoutCopay: allPlansWithoutCopay,
        ageBasedPricingCopay: allAgeBasedPricingCopay,
        ageBasedPricingNoCopay: allAgeBasedPricingNoCopay,
      });

      toast({
        title: "Arquivo processado com sucesso!",
        description: "Os dados foram extraídos e formatados.",
      });
    } catch (error) {
      console.error("Error parsing file:", error);
      toast({
        title: "Erro ao processar arquivo",
        description: "Verifique se o arquivo está no formato correto.",
        variant: "destructive",
      });
    }
  };

  const handleContinueWithInvalidFile = () => {
    if (!pendingWorkbook || !pendingFile) return;

    const workbook = pendingWorkbook;
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    const parseCurrency = (value: any): number => {
      if (typeof value === 'number') return value;
      if (typeof value === 'string') {
        const cleaned = value.replace(/[R$\\s.]/g, '').replace(',', '.');
        const parsed = parseFloat(cleaned);
        return isNaN(parsed) ? 0 : parsed;
      }
      return 0;
    };

    const companyName = (jsonData[1] as any)?.[1] || "";
    const concessionaire = (jsonData[2] as any)?.[1] || "";
    const broker = (jsonData[3] as any)?.[1] || "";

    const allDemographics: any[] = [];
    for (let i = 7; i <= 16; i++) {
      const row = jsonData[i] as any[];
      if (!row || !row[1]) continue;
      const ageRange = String(row[1]).trim();
      if (ageRange === "TOTAL" || ageRange === "IDADE MÉDIA:" || ageRange === "") continue;
      allDemographics.push({
        ageRange,
        titularM: row[2] || 0,
        titularF: row[3] || 0,
        dependentM: row[4] || 0,
        dependentF: row[5] || 0,
        agregadoM: 0,
        agregadoF: 0,
        totalM: row[8] || 0,
        totalF: row[9] || 0,
        total: row[10] || 0,
        percentage: row[11] || "0%",
      });
    }

    const allPlansWithCopay: any[] = [];
    for (let i = 26; i <= 33; i++) {
      const row = jsonData[i] as any[];
      if (!row || !row[1]) continue;
      const planName = String(row[1]).trim();
      if (!planName.startsWith("KLINI")) continue;
      allPlansWithCopay.push({
        name: planName,
        ansCode: String(row[4] || '').trim(),
        perCapita: parseCurrency(row[5]) || 0,
        estimatedInvoice: parseCurrency(row[6]) || 0,
      });
    }

    const copayAgeHeader = jsonData[40] as any[];
    const copayPlanNames = copayAgeHeader.slice(2).filter((name: string) => name && String(name).trim() !== "");
    const allAgeBasedPricingCopay: any[] = [];
    for (let i = 41; i <= 50; i++) {
      const row = jsonData[i] as any[];
      if (!row || !row[1]) continue;
      const ageRange = String(row[1]).trim();
      if (!ageRange.match(/\d+\s*-\s*\d+|59\+/)) continue;
      const pricing: any = { ageRange };
      copayPlanNames.forEach((planName: string, idx: number) => {
        pricing[planName] = parseCurrency(row[idx + 2]) || 0;
      });
      allAgeBasedPricingCopay.push(pricing);
    }

    const allPlansWithoutCopay: any[] = [];
    for (let i = 60; i <= 67; i++) {
      const row = jsonData[i] as any[];
      if (!row || !row[1]) continue;
      const planName = String(row[1]).trim();
      if (!planName.startsWith("KLINI")) continue;
      allPlansWithoutCopay.push({
        name: planName,
        ansCode: String(row[4] || '').trim(),
        perCapita: parseCurrency(row[5]) || 0,
        estimatedInvoice: parseCurrency(row[6]) || 0,
      });
    }

    const noCopayAgeHeader = jsonData[75] as any[];
    const noCopayPlanNames = noCopayAgeHeader.slice(2).filter((name: string) => name && String(name).trim() !== "");
    const allAgeBasedPricingNoCopay: any[] = [];
    for (let i = 76; i <= 85; i++) {
      const row = jsonData[i] as any[];
      if (!row || !row[1]) continue;
      const ageRange = String(row[1]).trim();
      if (!ageRange.match(/\d+\s*-\s*\d+|59\+/)) continue;
      const pricing: any = { ageRange };
      noCopayPlanNames.forEach((planName: string, idx: number) => {
        pricing[planName] = parseCurrency(row[idx + 2]) || 0;
      });
      allAgeBasedPricingNoCopay.push(pricing);
    }

    const emissionDate = new Date().toLocaleDateString('pt-BR');
    const validityDate = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toLocaleDateString('pt-BR');

    setParsedData({
      companyName,
      concessionaire,
      broker,
      emissionDate,
      validityDate,
      demographics: allDemographics,
      plansWithCopay: allPlansWithCopay,
      plansWithoutCopay: allPlansWithoutCopay,
      ageBasedPricingCopay: allAgeBasedPricingCopay,
      ageBasedPricingNoCopay: allAgeBasedPricingNoCopay,
    });

    setShowValidationDialog(false);
    setValidationError(null);
    setPendingFile(null);
    setPendingWorkbook(null);

    toast({
      title: "Arquivo processado!",
      description: "Os dados foram extraídos mesmo com avisos.",
    });
  };
  const handleExportPPTX = async () => {
    if (!parsedData) return;

    try {
      toast({
        title: "Gerando PowerPoint...",
        description: "Aguarde enquanto preparamos sua apresentação.",
      });

      await exportToPPTX(parsedData, coverImage);

      toast({
        title: "PowerPoint gerado com sucesso!",
        description: "Sua proposta foi exportada.",
      });
    } catch (error) {
      console.error("Error generating PPTX:", error);
      toast({
        title: "Erro ao gerar PowerPoint",
        description: "Tente novamente mais tarde.",
        variant: "destructive",
      });
    }
  };

  const handleExportPDF = async () => {
    const proposalElement = document.getElementById('proposal-content');
    if (!proposalElement) return;

    try {
      toast({
        title: "Gerando PDF...",
        description: "Aguarde enquanto preparamos sua proposta.",
      });

      const canvas = await html2canvas(proposalElement, {
        scale: 2,
        useCORS: true,
        logging: false,
      });

      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: 'a4',
      });

      const imgWidth = 210;
      const imgHeight = (canvas.height * imgWidth) / canvas.width;
      let heightLeft = imgHeight;
      let position = 0;

      pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
      heightLeft -= 297;

      while (heightLeft > 0) {
        position = heightLeft - imgHeight;
        pdf.addPage();
        pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
        heightLeft -= 297;
      }

      pdf.save(`Proposta_${parsedData?.companyName || 'Klini_Saude'}.pdf`);

      toast({
        title: "PDF gerado com sucesso!",
        description: "Sua proposta foi exportada.",
      });
    } catch (error) {
      console.error("Error generating PDF:", error);
      toast({
        title: "Erro ao gerar PDF",
        description: "Tente novamente mais tarde.",
        variant: "destructive",
      });
    }
  };

  const handleCoverImageChange = (imageDataUrl: string | null) => {
    setCoverImage(imageDataUrl);
    if (imageDataUrl) {
      toast({
        title: "Imagem da capa carregada!",
        description: "A capa personalizada será usada no PowerPoint.",
      });
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-[#1D7874] via-[#1a6b67] to-[#164e4b]">
      <div className="container mx-auto px-4 py-8">
        <div className="max-w-6xl mx-auto space-y-8">
          {/* Header Section */}
          <div className="text-center space-y-6 py-8">
            <div className="flex justify-center mb-6">
              <div className="bg-white/10 backdrop-blur-lg rounded-2xl p-6 border border-white/20">
                <img 
                  src={KLINI_LOGO}
                  alt="Klini Logo" 
                  className="h-20 w-auto"
                />
              </div>
            </div>
            <h1 className="text-5xl font-bold text-white drop-shadow-lg">
              Sistema de Cotação Klini Saúde
            </h1>
            <p className="text-xl text-white/90 font-light max-w-2xl mx-auto">
              Transforme suas planilhas em propostas profissionais com apenas alguns cliques
            </p>
          </div>

          {/* Main Upload Card */}
          <Card className="backdrop-blur-xl bg-white/95 border-none shadow-2xl">
            <div className="p-8 space-y-8">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                {/* Excel Upload */}
                <div className="space-y-3">
                  <label className="block text-sm font-semibold text-gray-700 uppercase tracking-wide">
                    📊 Planilha de Cotação
                  </label>
                  <input
                    id="file-upload"
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={handleFileUpload}
                    className="hidden"
                  />
                  <Button
                    onClick={() => document.getElementById('file-upload')?.click()}
                    variant="outline"
                    className="w-full h-32 border-2 border-dashed border-[#1D7874] hover:border-[#F7931E] hover:bg-[#FFF8F0] transition-all duration-300 group"
                  >
                    <div className="flex flex-col items-center gap-3">
                      <div className="p-3 bg-[#1D7874] group-hover:bg-[#F7931E] rounded-full transition-colors duration-300">
                        <Upload className="h-8 w-8 text-white" />
                      </div>
                      <div className="space-y-1">
                        <p className="font-semibold text-gray-700">Clique para fazer upload</p>
                        <p className="text-xs text-gray-500">Arquivos Excel (.xlsx ou .xls)</p>
                      </div>
                    </div>
                  </Button>
                </div>

                {/* Cover Upload */}
                <div className="space-y-3">
                  <label className="block text-sm font-semibold text-gray-700 uppercase tracking-wide">
                    🎨 Capa da Proposta
                  </label>
                  <CoverImageUpload onImageChange={handleCoverImageChange} currentImage={coverImage} />
                </div>
              </div>

              {/* Export Buttons */}
              {parsedData && (
                <div className="pt-6 border-t border-gray-200">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <Button
                      onClick={handleExportPDF}
                      className="h-14 bg-gradient-to-r from-[#1D7874] to-[#164e4b] hover:from-[#164e4b] hover:to-[#1D7874] text-white shadow-lg hover:shadow-xl transition-all duration-300 text-lg font-semibold"
                      size="lg"
                    >
                      <Download className="h-6 w-6 mr-3" />
                      Exportar para PDF
                    </Button>
                    <Button
                      onClick={handleExportPPTX}
                      className="h-14 bg-gradient-to-r from-[#F7931E] to-[#e67e0a] hover:from-[#e67e0a] hover:to-[#F7931E] text-white shadow-lg hover:shadow-xl transition-all duration-300 text-lg font-semibold"
                      size="lg"
                    >
                      <Presentation className="h-6 w-6 mr-3" />
                      Exportar para PowerPoint
                    </Button>
                  </div>
                </div>
              )}
            </div>
          </Card>

          {/* Preview Section */}
          {parsedData && (
            <div id="proposal-content" className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-700">
              <div className="text-center py-4">
                <h2 className="text-3xl font-bold text-white drop-shadow-lg">
                  📋 Pré-visualização da Proposta
                </h2>
              </div>

              <Card className="backdrop-blur-xl bg-white/95 border-none shadow-2xl overflow-hidden">
                <CompanyHeader
                  companyName={parsedData.companyName}
                  concessionaire={parsedData.concessionaire}
                  broker={parsedData.broker}
                  emissionDate={parsedData.emissionDate}
                  validityDate={parsedData.validityDate}
                />
              </Card>

              <Card className="backdrop-blur-xl bg-white/95 border-none shadow-2xl p-8">
                <h2 className="text-2xl font-bold mb-6 text-[#1D7874] flex items-center gap-3">
                  <span className="text-3xl">👥</span> Demografia
                </h2>
                <DemographicsTable data={parsedData.demographics} />
              </Card>

              {parsedData.plansWithCopay.length > 0 && (
                <Card className="backdrop-blur-xl bg-white/95 border-none shadow-2xl p-8">
                  <PricingTable 
                    title="Planos com Coparticipação"
                    plans={parsedData.plansWithCopay}
                  />
                </Card>
              )}

              {parsedData.plansWithoutCopay.length > 0 && (
                <Card className="backdrop-blur-xl bg-white/95 border-none shadow-2xl p-8">
                  <PricingTable 
                    title="Planos sem Coparticipação"
                    plans={parsedData.plansWithoutCopay}
                  />
                </Card>
              )}

              {parsedData.ageBasedPricingCopay.length > 0 && (
                <Card className="backdrop-blur-xl bg-white/95 border-none shadow-2xl p-8">
                  <AgeBasedPricingTable 
                    data={parsedData.ageBasedPricingCopay}
                    title="Valores por Faixa Etária - COM Coparticipação"
                  />
                </Card>
              )}

              {parsedData.ageBasedPricingNoCopay.length > 0 && (
                <Card className="backdrop-blur-xl bg-white/95 border-none shadow-2xl p-8">
                  <AgeBasedPricingTable 
                    data={parsedData.ageBasedPricingNoCopay}
                    title="Valores por Faixa Etária - SEM Coparticipação"
                  />
                </Card>
              )}

              <Card className="backdrop-blur-xl bg-white/95 border-none shadow-2xl p-8 text-center">
                <p className="text-sm text-gray-600 leading-relaxed">
                  Esta proposta foi elaborada levando em consideração as informações fornecidas através
                  do formulário de cotação enviado pela Corretora. No caso de implantação do contrato,
                  qualquer incompatibilidade implicará na inviabilidade ou reanálise da proposta.
                </p>
                <p className="text-xs text-gray-500 mt-4 font-semibold">ANS - Nº 42.202-9</p>
              </Card>
            </div>
          )}
        </div>
      </div>
      {/* Diálogo de Validação */}
   
    </div>
  );
};

export default Index;
