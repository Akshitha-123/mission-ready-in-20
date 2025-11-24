import { CheckCircle2, Circle } from "lucide-react";
import { cn } from "@/lib/utils";

interface Step {
  id: number;
  label: string;
  status: "complete" | "current" | "upcoming";
}

interface ProgressTrackerProps {
  currentStep: number;
}

const ProgressTracker = ({ currentStep }: ProgressTrackerProps) => {
  const steps: Step[] = [
    { id: 1, label: "Upload", status: currentStep > 1 ? "complete" : currentStep === 1 ? "current" : "upcoming" },
    { id: 2, label: "AI Draft", status: currentStep > 2 ? "complete" : currentStep === 2 ? "current" : "upcoming" },
    { id: 3, label: "Human Review", status: currentStep > 3 ? "complete" : currentStep === 3 ? "current" : "upcoming" },
    { id: 4, label: "Export", status: currentStep === 4 ? "current" : currentStep > 4 ? "complete" : "upcoming" },
  ];

  return (
    <div className="space-y-6">
      {steps.map((step, index) => (
        <div key={step.id} className="relative flex items-start gap-3">
          <div className="mt-0.5">
            {step.status === "complete" ? (
              <CheckCircle2 className="h-6 w-6 text-primary" />
            ) : step.status === "current" ? (
              <div className="relative">
                <Circle className="h-6 w-6 text-primary fill-primary" />
                <div className="absolute inset-0 animate-ping">
                  <Circle className="h-6 w-6 text-primary opacity-75" />
                </div>
              </div>
            ) : (
              <Circle className="h-6 w-6 text-muted-foreground" />
            )}
          </div>
          <span
            className={cn(
              "text-sm font-semibold",
              step.status === "current" && "text-primary",
              step.status === "complete" && "text-foreground",
              step.status === "upcoming" && "text-muted-foreground"
            )}
          >
            Step {step.id}: {step.label}
          </span>
          {index < steps.length - 1 && (
            <div className="absolute left-[11px] top-6 h-8 w-0.5 bg-border" />
          )}
        </div>
      ))}
    </div>
  );
};

export default ProgressTracker;
