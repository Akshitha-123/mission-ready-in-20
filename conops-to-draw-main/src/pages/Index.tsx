import { useEffect, useRef, useState } from "react";
import { toast } from "sonner";
import { Card } from "@/components/ui/card";
import ProgressTracker from "@/components/ProgressTracker";
import FileUploadZone from "@/components/FileUploadZone";
import ConopsViewer from "@/components/ConopsViewer";
import DrawDraftPanel, { DrawStatus } from "@/components/DrawDraftPanel";
import ActionBar from "@/components/ActionBar";
import { convertConopToPdf, generateDraw, uploadConop } from "@/lib/api";

const Index = () => {
  const [currentStep, setCurrentStep] = useState(1);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [conopsPreviewUrl, setConopsPreviewUrl] = useState<string | null>(null);
  const [conopsPreviewType, setConopsPreviewType] = useState<"pdf" | "pptx" | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isGeneratingPdf, setIsGeneratingPdf] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [drawStatus, setDrawStatus] = useState<DrawStatus>("idle");
  const [drawError, setDrawError] = useState<string | null>(null);
  const [drawPdfUrl, setDrawPdfUrl] = useState<string | null>(null);
  const [drawPreviewPdfUrl, setDrawPreviewPdfUrl] = useState<string | null>(null);
  const localPreviewRef = useRef<string | null>(null);
  const conversionJobRef = useRef(0);

  const revokeLocalPreview = () => {
    if (localPreviewRef.current) {
      URL.revokeObjectURL(localPreviewRef.current);
      localPreviewRef.current = null;
    }
  };

  const triggerPdfConversion = async (storedPath: string) => {
    const jobId = ++conversionJobRef.current;
    setIsGeneratingPdf(true);
    try {
      const conversion = await convertConopToPdf(storedPath);
      if (jobId !== conversionJobRef.current) {
        return;
      }
      setConopsPreviewUrl(conversion.preview_url);
      setConopsPreviewType("pdf");
      toast.success("PDF preview ready");
    } catch (error) {
      if (jobId === conversionJobRef.current) {
        const message = error instanceof Error ? error.message : "Unable to generate PDF preview";
        toast.error(message);
      }
    } finally {
      if (jobId === conversionJobRef.current) {
        setIsGeneratingPdf(false);
      }
    }
  };

  const handleFileSelect = async (file: File) => {
    revokeLocalPreview();
    conversionJobRef.current += 1;
    setIsGeneratingPdf(false);
    setConopsPreviewUrl(null);
    setConopsPreviewType(null);
    setSelectedFile(file);
    setCurrentStep(2);
    setIsProcessing(true);
    setIsUploading(true);
    setDrawPdfUrl(null);
    setDrawPreviewPdfUrl(null);
    setDrawStatus("parsing");
    setDrawError(null);

    try {
      const result = await uploadConop(file);
      if (result.preview_url) {
        setConopsPreviewUrl(result.preview_url);
        setConopsPreviewType("pdf");
      } else {
        const fallbackUrl = URL.createObjectURL(file);
        localPreviewRef.current = fallbackUrl;
        setConopsPreviewUrl(fallbackUrl);
        setConopsPreviewType("pptx");
        if (result.stored_path) {
          triggerPdfConversion(result.stored_path);
        }
      }
      setDrawStatus("generating");
      try {
        const drawResult = await generateDraw({
          filename: result.filename,
          raw_text: result.raw_text,
          sections: result.sections,
        });

        if (drawResult.draw && drawResult.draw_pdf_url) {
          setDrawPdfUrl(drawResult.draw_pdf_url);
          setDrawPreviewPdfUrl(drawResult.draw_pdf_preview_url ?? null);
          setDrawStatus("ready");
          setCurrentStep(3);
          toast.success("AI-generated DRAW is ready");
        } else {
          const errorMessage = drawResult.draw_error ?? "Unable to generate DRAW";
          setDrawPdfUrl(null);
          setDrawPreviewPdfUrl(null);
          setDrawStatus("error");
          setDrawError(errorMessage);
          setCurrentStep(2);
          toast.error(errorMessage);
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : "Failed to generate DRAW";
        setDrawPdfUrl(null);
        setDrawPreviewPdfUrl(null);
        setDrawStatus("error");
        setDrawError(message);
        setCurrentStep(2);
        toast.error(message);
      }
      toast.success("CONOP uploaded successfully");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Failed to upload CONOP";
      toast.error(message);
      setDrawStatus("error");
      setDrawError(message);
      setDrawPdfUrl(null);
      setDrawPreviewPdfUrl(null);
    } finally {
      setIsProcessing(false);
      setIsUploading(false);
    }
  };

  useEffect(() => {
    return () => {
      revokeLocalPreview();
    };
  }, []);

  const handleSave = () => {
    toast.success("Draft saved successfully");
    setCurrentStep(4);
  };

  const handleExport = async () => {
    if (!drawPdfUrl) return;

    try {
      const response = await fetch(drawPdfUrl);
      if (!response.ok) {
        throw new Error("Unable to download DRAW PDF");
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `DRAW_${new Date().toISOString().split("T")[0]}.pdf`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      toast.success("DRAW PDF downloaded");
      setCurrentStep(4);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Failed to download DRAW PDF";
      toast.error(message);
    }
  };

  return (
    <div className="min-h-screen w-screen flex flex-col bg-background">
      {/* Header */}
      <header className="bg-gradient-tactical border-b border-border shadow-tactical">
        <div className="px-8 lg:px-12 py-4">
          <div className="flex items-center gap-3">
            <div className="h-10 w-10 rounded bg-primary flex items-center justify-center font-bold text-primary-foreground shadow-glow">
              M20
            </div>
            <div>
              <h1 className="text-2xl font-bold text-foreground">Mission Ready in 20</h1>
              <p className="text-sm text-muted-foreground">CONOP to DRAW Conversion System</p>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 w-full px-6 sm:px-8 lg:px-12 py-10 flex flex-col">
        <div className="grid w-full gap-6 lg:grid-cols-[minmax(300px,1.05fr)_minmax(0,2.05fr)_minmax(0,2.05fr)] xl:grid-cols-[minmax(320px,1.1fr)_minmax(0,2.15fr)_minmax(0,2.15fr)] auto-rows-[minmax(0,1fr)] flex-1">
          <Card className="bg-card border-border shadow-tactical h-full">
            <div className="p-6 h-full flex flex-col">
              <div className="pb-4 border-b border-border">
                <p className="text-xs uppercase tracking-wide text-muted-foreground">Upload & Progress</p>
                <h2 className="text-lg font-bold text-foreground mt-1">Prepare CONOP</h2>
              </div>
              <div className="mt-6">
                <FileUploadZone onFileSelect={handleFileSelect} />
              </div>
              <div className="mt-8 flex-1 flex flex-col">
                <p className="text-xs uppercase tracking-wide text-muted-foreground">Workflow</p>
                <div className="mt-4 flex-1">
                  <ProgressTracker currentStep={currentStep} />
                </div>
              </div>
            </div>
          </Card>

          <div className="min-h-[60vh] h-full">
            <ConopsViewer
              previewUrl={conopsPreviewUrl}
              previewType={conopsPreviewType}
              isGeneratingPdf={isGeneratingPdf}
              isUploading={isUploading}
              fileName={selectedFile?.name || null}
            />
          </div>
          <div className="min-h-[60vh] h-full">
            <DrawDraftPanel
              pdfUrl={drawPdfUrl}
              previewPdfUrl={drawPreviewPdfUrl}
              status={drawStatus}
              errorMessage={drawError}
              fileName={selectedFile?.name || null}
              isPreviewReady={!!conopsPreviewUrl}
            />
          </div>
        </div>

        <div className="mt-10 -mx-6 sm:-mx-8 lg:-mx-12">
          <ActionBar
            hasDraft={!!drawPdfUrl && drawStatus === "ready"}
            currentStep={currentStep}
            onSave={handleSave}
            onExport={handleExport}
          />
        </div>
      </main>
    </div>
  );
};

export default Index;
