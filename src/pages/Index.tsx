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
import { LogoUpload } from "@/components/LogoUpload";
import { exportToPPTX } from "@/utils/pptxExport";

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
  const [logo, setLogo] = useState<string | null>(null);
  const { toast } = useToast();

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

      // Parse company info
      const companyName = (jsonData[0] as any)?.[1] || "";
      const concessionaire = (jsonData[1] as any)?.[1] || "";
      const broker = (jsonData[2] as any)?.[1] || "";

      // Parse demographics (rows 6-16)
      const allDemographics: any[] = [];
      for (let i = 6; i <= 16; i++) {
        const row = jsonData[i] as any[];
        if (!row || !row[0]) continue;
        const ageRange = String(row[0]).trim();
        if (ageRange === "TOTAL" || ageRange === "IDADE MÉDIA:" || ageRange === "") continue;
        allDemographics.push({
          ageRange,
          titularM: row[1] || 0,
          titularF: row[2] || 0,
          dependentM: row[3] || 0,
          dependentF: row[4] || 0,
          agregadoM: row[5] || 0,
          agregadoF: row[6] || 0,
          totalM: row[7] || 0,
          totalF: row[8] || 0,
          total: row[9] || 0,
          percentage: row[10] || "0%",
        });
      }

      // Parse plans WITH copay (rows 26-33, array index starts at 0)
      const allPlansWithCopay: any[] = [];
      for (let i = 26; i <= 33; i++) {
        const row = jsonData[i] as any[];
        if (!row || !row[0]) continue;
        const planName = String(row[0]).trim();
        if (!planName.startsWith("KLINI")) continue;
        allPlansWithCopay.push({
          name: planName,
          ansCode: String(row[3] || '').trim(),
          perCapita: parseCurrency(row[4]) || 0,
          estimatedInvoice: parseCurrency(row[5]) || 0,
        });
      }

      // Parse age-based pricing WITH copay (rows 42-51, array index starts at 0)
      const copayAgeHeader = jsonData[39] as any[];
      const copayPlanNames = copayAgeHeader.slice(1).filter((name: string) => name && String(name).trim() !== "");
      const allAgeBasedPricingCopay: any[] = [];
      for (let i = 42; i <= 51; i++) {
        const row = jsonData[i] as any[];
        if (!row || !row[0]) continue;
        const ageRange = String(row[0]).trim();
        if (!ageRange.match(/\d+\s*-\s*\d+|59\+/)) continue;
        const pricing: any = { ageRange };
        copayPlanNames.forEach((planName: string, idx: number) => {
          pricing[planName] = parseCurrency(row[idx + 1]) || 0;
        });
        allAgeBasedPricingCopay.push(pricing);
      }

      // Parse plans WITHOUT copay (rows 60-67, array index starts at 0)
      const allPlansWithoutCopay: any[] = [];
      for (let i = 60; i <= 67; i++) {
        const row = jsonData[i] as any[];
        if (!row || !row[0]) continue;
        const planName = String(row[0]).trim();
        if (!planName.startsWith("KLINI")) continue;
        allPlansWithoutCopay.push({
          name: planName,
          ansCode: String(row[3] || '').trim(),
          perCapita: parseCurrency(row[4]) || 0,
          estimatedInvoice: parseCurrency(row[5]) || 0,
        });
      }

      // Parse age-based pricing WITHOUT copay (rows 77-86, array index starts at 0)
      const noCopayAgeHeader = jsonData[74] as any[];
      const noCopayPlanNames = noCopayAgeHeader.slice(1).filter((name: string) => name && String(name).trim() !== "");
      const allAgeBasedPricingNoCopay: any[] = [];
      for (let i = 77; i <= 86; i++) {
        const row = jsonData[i] as any[];
        if (!row || !row[0]) continue;
        const ageRange = String(row[0]).trim();
        if (!ageRange.match(/\d+\s*-\s*\d+|59\+/)) continue;
        const pricing: any = { ageRange };
        noCopayPlanNames.forEach((planName: string, idx: number) => {
          pricing[planName] = parseCurrency(row[idx + 1]) || 0;
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

  const handleExportPPTX = async () => {
    if (!parsedData) return;

    try {
      toast({
        title: "Gerando PowerPoint...",
        description: "Aguarde enquanto preparamos sua apresentação.",
      });

      await exportToPPTX(parsedData);

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

  const handleLogoChange = (logoDataUrl: string | null) => {
    setLogo(logoDataUrl);
    if (logoDataUrl) {
      toast({
        title: "Logo carregada!",
        description: "A logo foi adicionada à sua proposta.",
      });
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-background via-background to-secondary/10">
      <div className="container mx-auto p-8">
        <div className="max-w-7xl mx-auto space-y-8">
          <div className="text-center space-y-4">
            <div className="flex justify-center mb-4">
              <img 
                src="/src/assets/logo-klini.webp" 
                alt="Klini Logo" 
                className="h-16 w-auto"
              />
            </div>
            <h1 className="text-4xl font-bold bg-gradient-to-r from-primary to-primary/60 bg-clip-text text-transparent">
              Sistema de Cotação Klini Saúde
            </h1>
            <p className="text-muted-foreground text-lg">
              Importe sua planilha e gere propostas profissionais em PDF ou PowerPoint
            </p>
          </div>

          <Card className="p-8 backdrop-blur-sm bg-card/50 border-2">
            <div className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label htmlFor="file-upload" className="block mb-2 text-sm font-medium">
                    Upload de Planilha Excel
                  </label>
                  <div className="relative">
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
                      className="w-full h-24 border-2 border-dashed hover:border-primary transition-colors"
                    >
                      <div className="flex flex-col items-center gap-2">
                        <Upload className="h-8 w-8" />
                        <span>Clique para fazer upload</span>
                        <span className="text-xs text-muted-foreground">
                          Arquivos .xlsx ou .xls
                        </span>
                      </div>
                    </Button>
                  </div>
                </div>

                <LogoUpload onLogoChange={handleLogoChange} currentLogo={logo} />
              </div>

              {parsedData && (
                <div className="flex gap-4 pt-4 border-t">
                  <Button
                    onClick={handleExportPDF}
                    className="flex-1 h-12 gap-2"
                    variant="default"
                  >
                    <Download className="h-5 w-5" />
                    Exportar para PDF
                  </Button>
                  <Button
                    onClick={handleExportPPTX}
                    className="flex-1 h-12 gap-2"
                    variant="secondary"
                  >
                    <Presentation className="h-5 w-5" />
                    Exportar para PowerPoint
                  </Button>
                </div>
              )}
            </div>
          </Card>

          {parsedData && (
            <div id="proposal-content" className="space-y-6">
              <CompanyHeader
                companyName={parsedData.companyName}
                concessionaire={parsedData.concessionaire}
                broker={parsedData.broker}
                emissionDate={parsedData.emissionDate}
                validityDate={parsedData.validityDate}
              />

              <Card className="p-6 backdrop-blur-sm bg-card/50">
                <h2 className="text-2xl font-bold mb-4 text-primary">Demografia</h2>
                <DemographicsTable data={parsedData.demographics} />
              </Card>

              {parsedData.plansWithCopay.length > 0 && (
                <PricingTable 
                  title="Planos com Coparticipação"
                  plans={parsedData.plansWithCopay}
                />
              )}

              {parsedData.plansWithoutCopay.length > 0 && (
                <PricingTable 
                  title="Planos sem Coparticipação"
                  plans={parsedData.plansWithoutCopay}
                />
              )}

              {parsedData.ageBasedPricingCopay.length > 0 && (
                <Card className="p-6 backdrop-blur-sm bg-card/50">
                  <AgeBasedPricingTable 
                    data={parsedData.ageBasedPricingCopay}
                    title="Valores por Faixa Etária - COM Coparticipação"
                  />
                </Card>
              )}

              {parsedData.ageBasedPricingNoCopay.length > 0 && (
                <Card className="p-6 backdrop-blur-sm bg-card/50">
                  <AgeBasedPricingTable 
                    data={parsedData.ageBasedPricingNoCopay}
                    title="Valores por Faixa Etária - SEM Coparticipação"
                  />
                </Card>
              )}

              <Card className="p-6 backdrop-blur-sm bg-card/50 text-center">
                <p className="text-sm text-muted-foreground">
                  Esta proposta foi elaborada levando em consideração as informações fornecidas através
                  do formulário de cotação enviado pela Corretora. No caso de implantação do contrato,
                  qualquer incompatibilidade implicará na inviabilidade ou reanálise da proposta.
                </p>
                <p className="text-xs text-muted-foreground mt-4">ANS - Nº 42.202-9</p>
              </Card>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default Index;
