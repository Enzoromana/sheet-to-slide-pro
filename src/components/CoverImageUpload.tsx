import { useState } from "react";
import { Upload, Eye, RefreshCw, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";

interface CoverImageUploadProps {
  onImageChange: (imageDataUrl: string | null) => void;
  currentImage: string | null;
}

export const CoverImageUpload = ({ onImageChange, currentImage }: CoverImageUploadProps) => {
  const [preview, setPreview] = useState<string | null>(currentImage);
  const [showFullView, setShowFullView] = useState(false);

  const handleImageUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!file.type.startsWith('image/')) {
      alert('Por favor, selecione apenas arquivos de imagem');
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      const result = e.target?.result as string;
      setPreview(result);
      onImageChange(result);
    };
    reader.readAsDataURL(file);
  };

  const handleRemoveImage = () => {
    setPreview(null);
    onImageChange(null);
    setShowFullView(false);
  };

  return (
    <>
      <div className="h-32 space-y-2">
        <input
          id="cover-upload"
          type="file"
          accept="image/*"
          onChange={handleImageUpload}
          className="hidden"
        />
        
        {preview ? (
          // Quando tem imagem - mostrar botões de ação
          <div className="h-full flex flex-col gap-2">
            <Button
              onClick={() => setShowFullView(true)}
              variant="outline"
              className="flex-1 border-2 border-[#1D7874] text-[#1D7874] hover:bg-[#1D7874] hover:text-white transition-all"
            >
              <Eye className="h-5 w-5 mr-2" />
              Visualizar Capa
            </Button>
            
            <div className="grid grid-cols-2 gap-2">
              <Button
                onClick={() => document.getElementById('cover-upload')?.click()}
                variant="outline"
                size="sm"
                className="border-[#1D7874] text-[#1D7874] hover:bg-[#1D7874] hover:text-white transition-all"
              >
                <RefreshCw className="h-4 w-4 mr-2" />
                Trocar
              </Button>
              <Button
                onClick={handleRemoveImage}
                variant="outline"
                size="sm"
                className="border-red-500 text-red-500 hover:bg-red-500 hover:text-white transition-all"
              >
                <X className="h-4 w-4 mr-2" />
                Remover
              </Button>
            </div>
          </div>
        ) : (
          // Quando não tem imagem - mostrar upload
          <Button
            onClick={() => document.getElementById('cover-upload')?.click()}
            variant="outline"
            className="w-full h-full border-2 border-dashed border-[#1D7874] hover:border-[#F7931E] hover:bg-[#FFF8F0] transition-all duration-300 group"
          >
            <div className="flex flex-col items-center gap-3">
              <div className="p-3 bg-[#1D7874] group-hover:bg-[#F7931E] rounded-full transition-colors duration-300">
                <Upload className="h-8 w-8 text-white" />
              </div>
              <div className="space-y-1">
                <p className="font-semibold text-gray-700">Upload de Capa</p>
                <p className="text-xs text-gray-500">PNG, JPG ou WEBP</p>
              </div>
            </div>
          </Button>
        )}
      </div>

      {/* Modal de Visualização Completa */}
      <Dialog open={showFullView} onOpenChange={setShowFullView}>
        <DialogContent className="max-w-4xl max-h-[90vh] p-0">
          <DialogHeader className="p-6 pb-0">
            <DialogTitle className="text-2xl font-bold text-[#1D7874]">
              📄 Visualização da Capa
            </DialogTitle>
          </DialogHeader>
          <div className="p-6 overflow-auto">
            {preview && (
              <img 
                src={preview} 
                alt="Capa completa" 
                className="w-full h-auto rounded-lg shadow-lg"
              />
            )}
          </div>
          <div className="p-6 pt-0 flex gap-3">
            <Button
              onClick={() => {
                setShowFullView(false);
                document.getElementById('cover-upload')?.click();
              }}
              className="flex-1 bg-[#1D7874] hover:bg-[#164e4b] text-white"
            >
              <RefreshCw className="h-4 w-4 mr-2" />
              Trocar Imagem
            </Button>
            <Button
              onClick={() => setShowFullView(false)}
              variant="outline"
              className="flex-1"
            >
              Fechar
            </Button>
          </div>
        </DialogContent>
      </Dialog>
    </>
  );
};
