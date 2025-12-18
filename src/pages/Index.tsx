import { useState } from "react";
import { Upload, Download, Presentation, X } from "lucide-react";
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
import { exportToPPTX } from "@/utils/pptxExport";
const COVER_DEC_2025 = null;

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
  // PRODUTOS G
  demographicsG?: any[];
  plansWithCopayG?: any[];
  plansWithoutCopayG?: any[];
  ageBasedPricingCopayG?: any[];
  ageBasedPricingNoCopayG?: any[];
}

const Index = () => {
  const [parsedData, setParsedData] = useState<ParsedData | null>(null);
  const [coverImage, setCoverImage] = useState<string | null>(null);
  const [includeProductosG, setIncludeProductosG] = useState(false);
  const { toast } = useToast();

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const parseCurrency = (value: any): number => {
      if (typeof value === "number") return value;
      if (typeof value === "string") {
        const cleaned = value.replace(/[R$\s.]/g, "").replace(",", ".");
        const parsed = parseFloat(cleaned);
        return isNaN(parsed) ? 0 : parsed;
      }
      return 0;
    };

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);
      
      // Parse ACIMA DE 100 VIDAS
      const sheet1 = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData1 = XLSX.utils.sheet_to_json(sheet1, { header: 1, defval: "" });

      const companyName = (jsonData1[1] as any)?.[1] || "";
      const concessionaire = (jsonData1[2] as any)?.[1] || "";
      const broker = (jsonData1[3] as any)?.[1] || "";

      // ACIMA DE 100 VIDAS - Demographics (rows 6-18 = indices 5-17)
      const allDemographics: any[] = [];
      for (let i = 5; i <= 17; i++) {
        const row = jsonData1[i] as any[];
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

      // ACIMA DE 100 VIDAS - Plans WITH copay (rows 26-34 = indices 25-33)
      const allPlansWithCopay: any[] = [];
      for (let i = 25; i <= 33; i++) {
        const row = jsonData1[i] as any[];
        if (!row || !row[1]) continue;
        const planName = String(row[1]).trim();
        if (!planName.startsWith("KLINI")) continue;
        allPlansWithCopay.push({
          name: planName,
          ansCode: String(row[4] || "").trim(),
          perCapita: parseCurrency(row[5]) || 0,
          estimatedInvoice: parseCurrency(row[6]) || 0,
        });
      }

      // ACIMA DE 100 VIDAS - Age pricing WITH copay (header row 39 = index 38, data rows 40-50 = indices 39-49)
      const copayAgeHeader = jsonData1[38] as any[];
      const copayPlanNames = copayAgeHeader.slice(2).filter((name: string) => name && String(name).trim() !== "");
      const allAgeBasedPricingCopay: any[] = [];
      for (let i = 39; i <= 49; i++) {
        const row = jsonData1[i] as any[];
        if (!row || !row[1]) continue;
        const ageRange = String(row[1]).trim();
        if (!ageRange.match(/\d+\s*-\s*\d+|59\+/)) continue;
        const pricing: any = { ageRange };
        copayPlanNames.forEach((planName: string, idx: number) => {
          pricing[planName] = parseCurrency(row[idx + 2]) || 0;
        });
        allAgeBasedPricingCopay.push(pricing);
      }

      // ACIMA DE 100 VIDAS - Plans WITHOUT copay (rows 58-66 = indices 57-65)
      const allPlansWithoutCopay: any[] = [];
      for (let i = 57; i <= 65; i++) {
        const row = jsonData1[i] as any[];
        if (!row || !row[1]) continue;
        const planName = String(row[1]).trim();
        if (!planName.startsWith("KLINI")) continue;
        allPlansWithoutCopay.push({
          name: planName,
          ansCode: String(row[4] || "").trim(),
          perCapita: parseCurrency(row[5]) || 0,
          estimatedInvoice: parseCurrency(row[6]) || 0,
        });
      }

      // ACIMA DE 100 VIDAS - Age pricing WITHOUT copay (header row 71 = index 70, data rows 72-82 = indices 71-81)
      const noCopayAgeHeader = jsonData1[70] as any[];
      const noCopayPlanNames = noCopayAgeHeader.slice(2).filter((name: string) => name && String(name).trim() !== "");
      const allAgeBasedPricingNoCopay: any[] = [];
      for (let i = 71; i <= 81; i++) {
        const row = jsonData1[i] as any[];
        if (!row || !row[1]) continue;
        const ageRange = String(row[1]).trim();
        if (!ageRange.match(/\d+\s*-\s*\d+|59\+/)) continue;
        const pricing: any = { ageRange };
        noCopayPlanNames.forEach((planName: string, idx: number) => {
          pricing[planName] = parseCurrency(row[idx + 2]) || 0;
        });
        allAgeBasedPricingNoCopay.push(pricing);
      }

      const emissionDate = new Date().toLocaleDateString("pt-BR");
      const validityDate = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toLocaleDateString("pt-BR");

      const dataToSet: ParsedData = {
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
      };

      // Parse PRODUTOS G if available
      if (workbook.SheetNames.includes("PRODUTOS G")) {
        const sheet2 = workbook.Sheets["PRODUTOS G"];
        const jsonData2 = XLSX.utils.sheet_to_json(sheet2, { header: 1, defval: "" });

        // PRODUTOS G - Demographics (rows 6-18 = indices 5-17)
        const allDemographicsG: any[] = [];
        for (let i = 5; i <= 17; i++) {
          const row = jsonData2[i] as any[];
          if (!row || !row[1]) continue;
          const ageRange = String(row[1]).trim();
          if (ageRange === "TOTAL" || ageRange === "IDADE MÉDIA:" || ageRange === "") continue;
          allDemographicsG.push({
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

        // PRODUTOS G - Plans WITH copay (rows 26-31 = indices 25-30)
        const allPlansWithCopayG: any[] = [];
        for (let i = 25; i <= 30; i++) {
          const row = jsonData2[i] as any[];
          if (!row || !row[1]) continue;
          const planName = String(row[1]).trim();
          if (!planName.startsWith("KLINI")) continue;
          allPlansWithCopayG.push({
            name: planName,
            ansCode: String(row[4] || "").trim(),
            perCapita: parseCurrency(row[5]) || 0,
            estimatedInvoice: parseCurrency(row[6]) || 0,
          });
        }

        // PRODUTOS G - Age pricing WITH copay (header row 35 = index 34, data rows 36-45 = indices 35-44)
        const copayAgeHeaderG = jsonData2[34] as any[];
        const copayPlanNamesG = copayAgeHeaderG.slice(2).filter((name: string) => name && String(name).trim() !== "");
        const allAgeBasedPricingCopayG: any[] = [];
        for (let i = 35; i <= 44; i++) {
          const row = jsonData2[i] as any[];
          if (!row || !row[1]) continue;
          const ageRange = String(row[1]).trim();
          if (!ageRange.match(/\d+\s*-\s*\d+|59\+/)) continue;
          const pricing: any = { ageRange };
          copayPlanNamesG.forEach((planName: string, idx: number) => {
            pricing[planName] = parseCurrency(row[idx + 2]) || 0;
          });
          allAgeBasedPricingCopayG.push(pricing);
        }

        // PRODUTOS G - Plans WITHOUT copay (rows 50+ - need to find exact range)
        const allPlansWithoutCopayG: any[] = [];
        for (let i = 49; i <= 55; i++) {
          const row = jsonData2[i] as any[];
          if (!row || !row[1]) continue;
          const planName = String(row[1]).trim();
          if (!planName.startsWith("KLINI")) continue;
          allPlansWithoutCopayG.push({
            name: planName,
            ansCode: String(row[4] || "").trim(),
            perCapita: parseCurrency(row[5]) || 0,
            estimatedInvoice: parseCurrency(row[6]) || 0,
          });
        }

        // PRODUTOS G - Age pricing WITHOUT copay
        const noCopayAgeHeaderG = jsonData2[59] as any[];
        const noCopayPlanNamesG = noCopayAgeHeaderG ? noCopayAgeHeaderG.slice(2).filter((name: string) => name && String(name).trim() !== "") : [];
        const allAgeBasedPricingNoCopayG: any[] = [];
        for (let i = 60; i <= 69; i++) {
          const row = jsonData2[i] as any[];
          if (!row || !row[1]) continue;
          const ageRange = String(row[1]).trim();
          if (!ageRange.match(/\d+\s*-\s*\d+|59\+/)) continue;
          const pricing: any = { ageRange };
          noCopayPlanNamesG.forEach((planName: string, idx: number) => {
            pricing[planName] = parseCurrency(row[idx + 2]) || 0;
          });
          allAgeBasedPricingNoCopayG.push(pricing);
        }

        dataToSet.demographicsG = allDemographicsG;
        dataToSet.plansWithCopayG = allPlansWithCopayG;
        dataToSet.plansWithoutCopayG = allPlansWithoutCopayG;
        dataToSet.ageBasedPricingCopayG = allAgeBasedPricingCopayG;
        dataToSet.ageBasedPricingNoCopayG = allAgeBasedPricingNoCopayG;
      }

      setParsedData(dataToSet);

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

      await exportToPPTX(parsedData, coverImage, includeProductosG);

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

  const handleCoverImageChange = (imageDataUrl: string | null) => {
    setCoverImage(imageDataUrl);
    if (imageDataUrl) {
      toast({
        title: "Imagem da capa carregada!",
        description: "A capa personalizada será usada no PowerPoint.",
      });
    }
  };

  const handleRemoveCover = () => {
    setCoverImage(COVER_DEC_2025);
    toast({
      title: "Capa removida!",
      description: "Usando capa padrão.",
    });
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-[#1D7874] via-[#1a6b67] to-[#164e4b]">
      <div className="container mx-auto px-4 py-8">
        <div className="max-w-6xl mx-auto space-y-8">
          {/* Header */}
          <div className="text-center space-y-3 py-6">
            <h1 className="text-4xl font-bold text-white drop-shadow-lg">Sistema de Cotação Klini Saúde</h1>
            <p className="text-lg text-white/90 font-light">Transforme suas planilhas em propostas profissionais com apenas alguns cliques</p>
          </div>

          {/* Main Content Card */}
          <Card className="backdrop-blur-xl bg-white/95 border-none shadow-2xl">
            <div className="p-8">
              {/* Grid 2 Colunas */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                {/* Coluna 1: Upload Planilha */}
                <div className="space-y-4">
                  <h2 className="text-lg font-bold text-[#1D7874] flex items-center gap-2">
                    📊 PLANILHA DE COTAÇÃO
                  </h2>
                  <input
                    id="file-upload"
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={handleFileUpload}
                    className="hidden"
                  />
                  <button
                    onClick={() => document.getElementById("file-upload")?.click()}
                    className="w-full h-32 border-2 border-dashed border-[#1D7874] rounded-lg hover:border-[#F7931E] hover:bg-[#FFF8F0] transition-all duration-300 flex flex-col items-center justify-center gap-3 cursor-pointer group"
                  >
                    <div className="p-3 bg-[#1D7874] group-hover:bg-[#F7931E] rounded-full transition-colors duration-300">
                      <Upload className="h-8 w-8 text-white" />
                    </div>
                    <div className="space-y-1 text-center">
                      <p className="font-semibold text-gray-700">Clique para fazer upload</p>
                      <p className="text-xs text-gray-500">Arquivos Excel (.xlsx ou .xls)</p>
                    </div>
                  </button>

                  {/* Checkbox para PRODUTOS G */}
                  {parsedData?.demographicsG && parsedData.demographicsG.length > 0 && (
                    <div className="flex items-center gap-3 p-4 bg-blue-50 rounded-lg border border-blue-200">
                      <input
                        type="checkbox"
                        id="produtos-g-checkbox"
                        checked={includeProductosG}
                        onChange={(e) => setIncludeProductosG(e.target.checked)}
                        className="w-5 h-5 text-[#1D7874] rounded cursor-pointer"
                      />
                      <label htmlFor="produtos-g-checkbox" className="cursor-pointer flex-1 font-semibold text-gray-700">
                        ✅ Incluir PRODUTOS G na apresentação
                      </label>
                    </div>
                  )}
                </div>

                {/* Coluna 2: Capa da Proposta */}
                <div className="space-y-4">
                  <h2 className="text-lg font-bold text-[#1D7874] flex items-center gap-2">
                    🎨 CAPA DA PROPOSTA
                  </h2>

                  {/* Preview da Capa */}
                  <div className="w-full h-48 bg-gradient-to-br from-[#1D7874] to-[#164e4b] rounded-lg border-2 border-gray-300 overflow-hidden flex items-center justify-center">
                    {coverImage ? (
                      <img src={coverImage} alt="Capa" className="w-full h-full object-cover" />
                    ) : (
                      <div className="text-white text-center">
                        <p className="text-sm">Sem capa selecionada</p>
                      </div>
                    )}
                  </div>

                  {/* Botões Trocar/Remover */}
                 <div className="grid grid-cols-2 gap-3">
  <button
    onClick={() => document.getElementById("cover-upload")?.click()}
    className="px-4 py-2 border-2 border-green-600 text-green-600 rounded-lg hover:bg-green-50 transition-colors duration-300 flex items-center justify-center gap-2 font-semibold text-sm"
  >
    <Upload className="h-4 w-4" />
    Trocar
  </button>
  <input
    id="cover-upload"
    type="file"
    accept="image/*"
    onChange={(e) => {
      const file = e.target.files?.[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (event) => {
          handleCoverImageChange(event.target?.result as string);
        };
        reader.readAsDataURL(file);
      }
    }}
    className="hidden"
  />
  <button
                      onClick={handleRemoveCover}
                      className="px-4 py-2 border-2 border-red-500 text-red-500 rounded-lg hover:bg-red-50 transition-colors duration-300 flex items-center justify-center gap-2 font-semibold text-sm"
                    >
                      <X className="h-4 w-4" />
                      Remover
                    </button>
                  </div>
                </div>
              </div>

              {/* Export Buttons */}
              {parsedData && (
                <div className="pt-8 border-t border-gray-200">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
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
        </div>
      </div>
    </div>
  );
};

export default Index;
