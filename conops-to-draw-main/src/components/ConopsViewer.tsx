import { useEffect, useRef, useState } from "react";
import { init as initPptxPreview, type PPTXPreviewer } from "pptx-preview";
import { FileText } from "lucide-react";
import { Card } from "@/components/ui/card";

interface ConopsViewerProps {
  previewUrl: string | null;
  fileName: string | null;
  previewType: "pdf" | "pptx" | null;
  isGeneratingPdf?: boolean;
  isUploading?: boolean;
}

const ConopsViewer = ({ previewUrl, fileName, previewType, isGeneratingPdf = false, isUploading = false }: ConopsViewerProps) => {
  const iframeRef = useRef<HTMLIFrameElement | null>(null);
  const containerRef = useRef<HTMLDivElement | null>(null);
  const previewerRef = useRef<PPTXPreviewer | null>(null);
  const [previewError, setPreviewError] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);

  useEffect(() => {
    if (previewType !== "pptx") {
      previewerRef.current?.destroy();
      previewerRef.current = null;
      if (containerRef.current) {
        containerRef.current.innerHTML = "";
      }
      return;
    }

    const container = containerRef.current;
    if (!previewUrl || !container) {
      setPreviewError(null);
      setIsLoading(false);
      return;
    }

    setIsLoading(true);
    setPreviewError(null);
    container.innerHTML = "";
    const instance = initPptxPreview(container, {
      mode: "list",
      width: container.clientWidth || undefined,
      height: container.clientHeight || undefined,
    });
    previewerRef.current = instance;
    let aborted = false;

    fetch(previewUrl)
      .then((response) => response.arrayBuffer())
      .then((buffer) => {
        if (aborted) return;
        return instance.preview(buffer);
      })
      .catch(() => {
        if (aborted) return;
        setPreviewError("Inline preview is unavailable in this browser.");
      })
      .finally(() => {
        if (!aborted) {
          setIsLoading(false);
        }
      });

    return () => {
      aborted = true;
      instance.destroy();
      previewerRef.current = null;
      container.innerHTML = "";
    };
  }, [previewType, previewUrl]);

  useEffect(() => {
    if (previewType === "pdf" && previewUrl) {
      setIsLoading(true);
      setPreviewError(null);
      return;
    }
    if (!previewUrl) {
      setIsLoading(false);
      setPreviewError(null);
    }
  }, [previewType, previewUrl]);

  return (
    <Card className="h-full bg-card border-border shadow-tactical">
      <div className="p-6 h-full flex flex-col">
        <div className="flex items-center gap-2 mb-4 pb-4 border-b border-border">
          <FileText className="h-5 w-5 text-secondary" />
          <h2 className="text-lg font-bold text-foreground">Uploaded CONOP</h2>
        </div>
        {previewUrl ? (
          <div className="flex-1 flex flex-col overflow-hidden">
            <div className="mb-3 text-sm text-muted-foreground">
              File: <span className="text-secondary font-semibold">{fileName}</span>
            </div>
            {previewType === "pdf" ? (
              <div className="flex-1 rounded border border-border overflow-hidden bg-background">
                <iframe
                  ref={iframeRef}
                  src={previewUrl ?? undefined}
                  title="CONOP preview"
                  className="w-full h-full"
                  onLoad={() => setIsLoading(false)}
                  onError={() => {
                    setIsLoading(false);
                    setPreviewError("Unable to load the generated PDF preview.");
                  }}
                />
              </div>
            ) : (
              <div
                className="flex-1 rounded border border-border overflow-hidden bg-background"
                ref={containerRef}
              />
            )}
            {isLoading && (
              <p className="mt-3 text-xs text-muted-foreground">Loading CONOP preview...</p>
            )}
            {isGeneratingPdf && !previewError && (
              <p className="mt-1 text-xs text-muted-foreground">Generating high-fidelity PDF preview…</p>
            )}
            {previewError && (
              <p className="mt-3 text-xs text-muted-foreground">
                {previewError} <span className="block">Please download the file and open it locally.</span>
              </p>
            )}
          </div>
        ) : (
          <div className="flex-1 flex flex-col items-center justify-center text-muted-foreground gap-3 text-center px-6">
            {isUploading ? (
              <>
                <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-secondary" />
                <p className="text-sm">Uploading CONOP… preparing preview.</p>
              </>
            ) : (
              <>
                <FileText className="h-16 w-16 opacity-50" />
                <p className="text-sm">Upload a CONOP file to view it here.</p>
              </>
            )}
          </div>
        )}
      </div>
    </Card>
  );
};

export default ConopsViewer;
