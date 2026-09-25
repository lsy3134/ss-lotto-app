import { Router, type IRouter } from "express";

const router: IRouter = Router();

router.post("/ocr-schedule", async (req, res) => {
  try {
    const apiKey = process.env.GOOGLE_CLOUD_VISION_API_KEY;
    if (!apiKey) {
      res.status(503).json({ error: "GOOGLE_CLOUD_VISION_API_KEY is not configured" });
      return;
    }
    const { imageBase64, mimeType } = req.body as { imageBase64?: string; mimeType?: string };
    if (!imageBase64 || !mimeType?.startsWith("image/")) {
      res.status(400).json({ error: "A valid imageBase64 and image mimeType are required" });
      return;
    }

    const response = await fetch(`https://vision.googleapis.com/v1/images:annotate?key=${encodeURIComponent(apiKey)}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        requests: [{ image: { content: imageBase64 }, features: [{ type: "DOCUMENT_TEXT_DETECTION" }] }],
      }),
    });
    const payload = await response.json() as {
      responses?: Array<{ fullTextAnnotation?: { text?: string }; textAnnotations?: Array<{ description?: string }>; error?: { message?: string } }>;
      error?: { message?: string };
    };
    if (!response.ok || payload.error || payload.responses?.[0]?.error) {
      res.status(502).json({ error: payload.error?.message ?? payload.responses?.[0]?.error?.message ?? "Vision OCR failed" });
      return;
    }
    const first = payload.responses?.[0];
    res.json({ text: first?.fullTextAnnotation?.text ?? first?.textAnnotations?.[0]?.description ?? "" });
  } catch (err) {
    console.error("Schedule OCR error:", err);
    res.status(500).json({ error: "OCR processing failed" });
  }
});

export default router;
