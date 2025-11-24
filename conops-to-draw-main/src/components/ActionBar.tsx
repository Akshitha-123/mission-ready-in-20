import { Download, Save } from "lucide-react";
import { Button } from "@/components/ui/button";

interface ActionBarProps {
  hasDraft: boolean;
  currentStep: number;
  onSave: () => void | Promise<void>;
  onExport: () => void | Promise<void>;
}

const statusByStep: Record<number, string> = {
  1: "Awaiting CONOP upload",
  2: "Processing CONOP and drafting DRAW",
  3: "Review and export DRAW",
  4: "DRAW finalized",
};

const ActionBar = ({ hasDraft, currentStep, onSave, onExport }: ActionBarProps) => {
  const statusMessage = statusByStep[currentStep] ?? "System ready";
  return (
    <div className="bg-card border-t border-border px-6 sm:px-8 lg:px-12 py-4">
      <div className="flex items-center justify-between">
        <div className="flex gap-3">
          <Button 
            variant="outline" 
            onClick={onSave}
            disabled={!hasDraft}
            className="gap-2"
          >
            <Save className="h-4 w-4" />
            Save Draft
          </Button>
          <Button 
            onClick={onExport}
            disabled={!hasDraft}
            className="gap-2 bg-primary hover:bg-primary/90 text-primary-foreground shadow-glow"
          >
            <Download className="h-4 w-4" />
            Export Final DRAW
          </Button>
        </div>
        <div className="flex items-center gap-2">
          <div className="h-2 w-2 rounded-full bg-primary animate-pulse"></div>
          <span className="text-sm font-semibold text-foreground">{statusMessage}</span>
        </div>
      </div>
    </div>
  );
};

export default ActionBar;
