import { Router } from "express";
import { GoogleGenAI } from "@google/genai";

const router = Router();

const ai = new GoogleGenAI({
  apiKey: process.env.AI_INTEGRATIONS_GEMINI_API_KEY ?? "dummy",
  httpOptions: {
    apiVersion: "",
    baseUrl: process.env.AI_INTEGRATIONS_GEMINI_BASE_URL,
  },
});

router.post("/ocr-calendar", async (req, res) => {
  try {
    const { imageBase64, mimeType } = req.body as {
      imageBase64: string;
      mimeType: string;
    };

    if (!imageBase64 || !mimeType) {
      res.status(400).json({ error: "imageBase64 and mimeType are required" });
      return;
    }

    const result = await ai.models.generateContent({
      model: "gemini-2.5-flash",
      contents: [
        {
          role: "user",
          parts: [
            {
              inlineData: {
                mimeType,
                data: imageBase64,
              },
            },
            {
              text: `이 이미지는 한국어로 작성된 월별 휴무 달력입니다.
각 날짜에 적힌 사람 이름들을 모두 추출해주세요.
달력 제목에 나온 월을 기준으로 날짜를 파악하세요.

반드시 다음 JSON 형식으로만 응답하세요 (설명 없이 JSON만):
{
  "1": ["이름1", "이름2"],
  "2": ["이름3"],
  "3": [],
  ...
}

- 키는 날짜 숫자(1~31)를 문자열로
- 값은 해당 날짜에 휴무인 사람 이름 배열
- 이름이 없는 날짜는 빈 배열 []
- JSON 외 다른 텍스트 절대 금지`,
            },
          ],
        },
      ],
    });

    const text = result.text ?? "";

    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      res.status(500).json({ error: "OCR 결과를 파싱할 수 없습니다", raw: text });
      return;
    }

    const parsed: Record<string, string[]> = JSON.parse(jsonMatch[0]);
    res.json({ dates: parsed });
  } catch (err) {
    console.error("OCR error:", err);
    res.status(500).json({ error: String(err) });
  }
});

export default router;
