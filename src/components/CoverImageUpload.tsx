import { useState } from "react";
import { Upload } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";

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
    <Card className="p-4">
      <label className="block mb-2 text-sm font-medium">
        Imagem da Capa (Opcional)
      </label>
      
      {preview ? (
        <div className="space-y-2">
          <div className="border-2 border-dashed rounded-lg p-2">
            <img 
              src={preview} 
              alt="Capa preview" 
              className="w-full h-32 object-cover rounded"
            />
          </div>
          <div className="flex gap-2">
            <Button
              onClick={() => document.getElementById('cover-upload')?.click()}
              variant="outline"
              size="sm"
              className="flex-1"
            >
              Trocar Imagem
            </Button>
            <Button
              onClick={handleRemoveImage}
              variant="destructive"
              size="sm"
              className="flex-1"
            >
              Remover
            </Button>
          </div>
        </div>
      ) : (
        <div className="relative">
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
            className="w-full h-24 border-2 border-dashed hover:border-primary transition-colors"
          >
            <div className="flex flex-col items-center gap-2">
              <Upload className="h-6 w-6" />
              <span className="text-xs">Upload da Capa</span>
              <span className="text-xs text-muted-foreground">
                PNG, JPG ou WEBP
              </span>
            </div>
          </Button>
        </div>
      )}
    </Card>
  );
};
