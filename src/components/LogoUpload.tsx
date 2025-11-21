import { useState } from "react";
import { Upload, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";

interface LogoUploadProps {
  onLogoChange: (logoUrl: string | null) => void;
  currentLogo: string | null;
}

export const LogoUpload = ({ onLogoChange, currentLogo }: LogoUploadProps) => {
  const [preview, setPreview] = useState<string | null>(currentLogo);

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        const result = reader.result as string;
        setPreview(result);
        onLogoChange(result);
      };
      reader.readAsDataURL(file);
    }
  };

  const handleRemoveLogo = () => {
    setPreview(null);
    onLogoChange(null);
  };

  return (
    <Card className="p-4 border-2 border-dashed border-primary/30">
      <div className="flex items-center gap-4">
        <div className="flex-1">
          <h4 className="font-semibold text-foreground mb-2">Logo da Empresa</h4>
          {preview ? (
            <div className="flex items-center gap-4">
              <img 
                src={preview} 
                alt="Logo preview" 
                className="h-16 object-contain rounded border border-border"
              />
              <Button
                variant="outline"
                size="sm"
                onClick={handleRemoveLogo}
                className="text-destructive hover:bg-destructive/10"
              >
                <X className="h-4 w-4 mr-1" />
                Remover
              </Button>
            </div>
          ) : (
            <label htmlFor="logo-upload">
              <Button variant="outline" size="sm" className="cursor-pointer" asChild>
                <span>
                  <Upload className="h-4 w-4 mr-2" />
                  Upload Logo
                </span>
              </Button>
              <input
                id="logo-upload"
                type="file"
                accept="image/*"
                className="hidden"
                onChange={handleFileChange}
              />
            </label>
          )}
        </div>
      </div>
    </Card>
  );
};
