import { useState } from "react";
import { Upload, X, RefreshCw } from "lucide-react";
import { Button } from "@/components/ui/button";

interface CoverImageUploadProps {
  onImageChange: (imageDataUrl: string | null) => void;
  currentImage: string | null;
}

export const CoverImageUpload = ({ onImageChange, currentImage }: CoverImageUploadProps) => {
  const [preview, setPreview] = useState<string | null>(currentImage);

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
  };

  return (
    <div className="h-full">
      {preview ? (
        <div className="space-y-3 h-full flex flex-col">
          <div className="border-2 border-[#1D7874]/20 rounded-xl overflow-hidden shadow-md flex-1">
            <img 
              src={preview} 
              alt="Capa preview" 
              className="w-full h-full object-cover"
            />
          </div>
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
        <div className="h-32">
          <input
            id="cover-upload"
            type="file"
            accept="image/*"
            onChange={handleImageUpload}
            className="hidden"
          />
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
        </div>
      )}
      <input
        id="cover-upload"
        type="file"
        accept="image/*"
        onChange={handleImageUpload}
        className="hidden"
      />
    </div>
  );
};
