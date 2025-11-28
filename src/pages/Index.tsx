// Copie TODO este arquivo e substitua seu Index.tsx

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

  return (
    <div className="min-h-screen bg-gradient-to-br from-teal-50 to-blue-50">
      <div className="container mx-auto px-4 py-8">
        <div className="text-center mb-12">
          <div className="flex items-center justify-center mb-4">
            <div dangerouslySetInnerHTML={{ __html: KLINI_LOGO }} className="w-12 h-12" />
          </div>
          <h1 className="text-4xl font-bold text-teal-900 mb-2">
            Conversor de Propostas Klini Saúde
          </h1>
          <p className="text-gray-600">
            Transforme Planilhas em Propostas Profissionais
          </p>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          <Card className="lg:col-span-1 p-6">
            <div className="space-y-6">
              <div>
                <h2 className="text-xl font-semibold text-teal-900 mb-4">
                  1. Importar Arquivo
                </h2>
                <label className="block border-2 border-dashed border-teal-300 rounded-lg p-8 text-center cursor-pointer hover:border-teal-500 transition">
                  <Upload className="w-12 h-12 text-teal-600 mx-auto mb-2" />
                  <p className="text-teal-700 font-semibold">
                    Clique para selecionar
                  </p>
                  <p className="text-sm text-gray-500">arquivo Excel (.xlsx)</p>
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={handleFileUpload}
                    className="hidden"
                  />
                </label>
              </div>

              {parsedData && (
                <>
                  <div>
                    <h2 className="text-xl font-semibold text-teal-900 mb-4">
                      2. Customizar Capa
                    </h2>
                    <CoverImageUpload onImageSelect={setCoverImage} />
                  </div>

                  <div>
                    <h2 className="text-xl font-semibold text-teal-900 mb-4">
                      3. Exportar Apresentação
                    </h2>
                    <Button
                      onClick={handleExportPPTX}
                      className="w-full bg-teal-600 hover:bg-teal-700 text-white rounded-lg py-3 font-semibold flex items-center justify-center gap-2"
                    >
                      <Presentation className="w-5 h-5" />
                      Gerar PowerPoint
                    </Button>
                  </div>
                </>
              )}
            </div>
          </Card>

          <div className="lg:col-span-2 space-y-6">
            {parsedData && (
              <>
                <CompanyHeader
                  companyName={parsedData.companyName}
                  concessionaire={parsedData.concessionaire}
                  broker={parsedData.broker}
                  emissionDate={parsedData.emissionDate}
                  validityDate={parsedData.validityDate}
                />

                <DemographicsTable data={parsedData.demographics} />

                <PricingTable
                  title="Planos com Coparticipação"
                  plans={parsedData.plansWithCopay}
                />

                <AgeBasedPricingTable
                  title="Valores por Faixa Etária - COM Coparticipação"
                  data={parsedData.ageBasedPricingCopay}
                />

                <PricingTable
                  title="Planos sem Coparticipação"
                  plans={parsedData.plansWithoutCopay}
                />

                <AgeBasedPricingTable
                  title="Valores por Faixa Etária - SEM Coparticipação"
                  data={parsedData.ageBasedPricingNoCopay}
                />
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default Index;
