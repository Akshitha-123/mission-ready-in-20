import { AlertTriangle, Clock, FileCheck, Loader2 } from "lucide-react";
import { Card } from "@/components/ui/card";

export type DrawStatus = "idle" | "parsing" | "generating" | "ready" | "error";

interface DrawDraftPanelProps {
  pdfUrl: string | null;
  previewPdfUrl?: string | null;
  status: DrawStatus;
  errorMessage?: string | null;
  fileName?: string | null;
  isPreviewReady: boolean;
}

const DrawDraftPanel = ({ pdfUrl, previewPdfUrl, status, errorMessage, fileName, isPreviewReady }: DrawDraftPanelProps) => {
  const isBusy = isPreviewReady && (status === "parsing" || status === "generating");
  const backendBase = import.meta.env.VITE_API_URL || "http://localhost:8000";

  const displayUrl = previewPdfUrl
    ? `${backendBase}${previewPdfUrl}`
    : pdfUrl
    ? `${backendBase}${pdfUrl}`
    : null;

  const busyCopy =
    status === "parsing"
      ? {
          title: "Parsing CONOP",
          body: "Extracting structure and sections from the uploaded briefing…",
        }
      : {
          title: "Building DRAW PDF",
          body: "Generating the DRAW…",
        };

  return (
    <Card className="h-full bg-card border-border shadow-tactical">
      <div className="p-6 h-full flex flex-col">
        <div className="flex items-center gap-2 mb-4 pb-4 border-b border-border">
          <FileCheck className="h-5 w-5 text-primary" />
          <h2 className="text-lg font-bold text-foreground">AI-Generated DRAW Draft</h2>
        </div>
        {isBusy ? (
          <div className="flex-1 flex flex-col items-center justify-center text-muted-foreground gap-2 text-center px-8">
            <Loader2 className="h-10 w-10 opacity-80 animate-spin" />
            <p className="text-base font-semibold text-foreground">{busyCopy.title}</p>
            <p className="text-sm">{busyCopy.body}</p>
          </div>
        ) : status === "error" ? (
          <div className="flex-1 flex flex-col items-center justify-center text-center text-destructive gap-3 px-6">
            <AlertTriangle className="h-10 w-10" />
            <div>
              <p className="font-semibold">Unable to generate DRAW</p>
              <p className="text-sm text-muted-foreground">{errorMessage ?? "Please try uploading again."}</p>
            </div>
          </div>
        ) : displayUrl && status === "ready" ? (
          <div className="flex-1 flex flex-col overflow-hidden">
            <div className="bg-gradient-success/10 border border-primary/20 rounded p-4 mb-3">
              <p className="text-primary text-sm font-semibold mb-1">✓ DRAW PDF Ready</p>
              <p className="text-xs text-muted-foreground">Review the generated DRAW below.</p>
              {fileName && (
                <p className="text-xs text-muted-foreground mt-1">
                  Source CONOP: <span className="text-foreground">{fileName}</span>
                </p>
              )}
            </div>
            <div className="flex-1 border border-border rounded overflow-hidden bg-background">
              <iframe
                title="DRAW PDF preview"
                src={displayUrl ? `${displayUrl}#toolbar=0&statusbar=0` : undefined}
                className="w-full h-full"
              />
            </div>
          </div>
        ) : (
          <div className="flex-1 flex flex-col items-center justify-center text-muted-foreground">
            <Clock className="h-16 w-16 mb-4 opacity-50" />
            <p className="text-sm text-center px-6">DRAW PDF will be generated after CONOP upload</p>
          </div>
        )}
      </div>
    </Card>
  );
};

export default DrawDraftPanel;
