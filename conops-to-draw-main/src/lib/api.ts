export interface UploadConopResponse {
  filename: string;
  stored_path: string;
  raw_text: string;
  sections: Record<string, string>;
  preview_url: string | null;
}

export interface GenerateDrawRequestPayload {
  filename: string;
  raw_text: string;
  sections: Record<string, string>;
}

export interface GenerateDrawResponse {
  draw: Record<string, unknown> | null;
  draw_error?: string | null;
  draw_pdf_url: string | null;
  draw_pdf_preview_url: string | null;
}

export interface ConvertPreviewResponse {
  preview_url: string;
}

const API_BASE_URL = (import.meta.env.VITE_API_URL ?? "").replace(/\/$/, "");
const buildApiUrl = (path: string) => (API_BASE_URL ? `${API_BASE_URL}${path}` : path);
const API_UPLOAD_URL = buildApiUrl("/api/conops/upload");
const API_CONVERT_URL = buildApiUrl("/api/conops/convert-preview");
const API_GENERATE_DRAW_URL = buildApiUrl("/api/conops/generate-draw");

async function parseJsonResponse<T>(response: Response, fallbackMessage: string): Promise<T> {
  const responseBody = await response.text();

  if (!response.ok) {
    let message = responseBody.trim() || fallbackMessage;
    try {
      const parsed = JSON.parse(responseBody);
      if (typeof parsed?.detail === "string" && parsed.detail.trim()) {
        message = parsed.detail.trim();
      }
    } catch {
      // response is not JSON; keep fallback message
    }
    throw new Error(message);
  }

  try {
    return JSON.parse(responseBody) as T;
  } catch {
    throw new Error("Unexpected response from CONOP API");
  }
}

export async function uploadConop(file: File): Promise<UploadConopResponse> {
  const formData = new FormData();
  formData.append("file", file);

  const response = await fetch(API_UPLOAD_URL, {
    method: "POST",
    body: formData,
  });
  const data = await parseJsonResponse<UploadConopResponse>(response, "Failed to upload CONOP");

  if (API_BASE_URL && data.preview_url && !data.preview_url.startsWith("http")) {
    data.preview_url = new URL(data.preview_url, `${API_BASE_URL}/`).toString();
  }
  return data;
}

export async function convertConopToPdf(storedPath: string): Promise<ConvertPreviewResponse> {
  const response = await fetch(API_CONVERT_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ stored_path: storedPath }),
  });
  const data = await parseJsonResponse<ConvertPreviewResponse>(response, "Failed to generate PDF preview");

  if (API_BASE_URL && data.preview_url && !data.preview_url.startsWith("http")) {
    data.preview_url = new URL(data.preview_url, `${API_BASE_URL}/`).toString();
  }
  return data;
}

export async function generateDraw(payload: GenerateDrawRequestPayload): Promise<GenerateDrawResponse> {
  const response = await fetch(API_GENERATE_DRAW_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });
  const data = await parseJsonResponse<GenerateDrawResponse>(response, "Failed to generate DRAW");

  if (API_BASE_URL && data.draw_pdf_url && !data.draw_pdf_url.startsWith("http")) {
    data.draw_pdf_url = new URL(data.draw_pdf_url, `${API_BASE_URL}/`).toString();
  }
  if (API_BASE_URL && data.draw_pdf_preview_url && !data.draw_pdf_preview_url.startsWith("http")) {
    data.draw_pdf_preview_url = new URL(data.draw_pdf_preview_url, `${API_BASE_URL}/`).toString();
  }
  return data;
}
