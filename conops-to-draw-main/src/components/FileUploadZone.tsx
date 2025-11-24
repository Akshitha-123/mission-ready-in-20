import { Upload } from "lucide-react";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";
import { useState } from "react";

interface FileUploadZoneProps {
  onFileSelect: (file: File) => void;
}

const FileUploadZone = ({ onFileSelect }: FileUploadZoneProps) => {
  const [isDragging, setIsDragging] = useState(false);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && file.name.toLowerCase().endsWith(".pptx")) {
      onFileSelect(file);
    }
  };

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      onFileSelect(file);
    }
  };

  return (
    <div
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      className={cn(
        "border-2 border-dashed rounded-lg p-8 text-center transition-all duration-200",
        isDragging ? "border-primary bg-primary/5 shadow-glow" : "border-border hover:border-primary/50"
      )}
    >
      <Upload className="h-12 w-12 mx-auto mb-4 text-muted-foreground" />
      <h3 className="text-lg font-bold mb-2 text-foreground">Upload CONOP</h3>
      <p className="text-sm text-muted-foreground mb-4">
        Drop CONOP file or click to browse<br />
        <span className="text-xs">(PowerPoint .pptx files only)</span>
      </p>
      <Button variant="outline" className="relative">
        <input
          type="file"
          accept=".pptx"
          onChange={handleFileInput}
          className="absolute inset-0 opacity-0 cursor-pointer"
        />
        Select File
      </Button>
    </div>
  );
};

export default FileUploadZone;
