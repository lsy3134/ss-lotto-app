import { Router, type IRouter } from "express";
import { db, holidayStoreTable } from "@workspace/db";
import { eq } from "drizzle-orm";

const router: IRouter = Router();

router.get("/holiday-map", async (_req, res) => {
  res.set("Cache-Control", "no-store, no-cache, must-revalidate");
  res.set("Pragma", "no-cache");
  res.set("Expires", "0");
  try {
    const rows = await db.select().from(holidayStoreTable).where(eq(holidayStoreTable.id, 1));
    if (rows.length === 0) {
      res.json({ fileName: "", holidayMap: {}, updatedAt: null });
      return;
    }
    res.json({
      fileName: rows[0].fileName,
      holidayMap: rows[0].holidayMap,
      updatedAt: rows[0].updatedAt ? rows[0].updatedAt.toISOString() : null,
    });
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

router.post("/holiday-map", async (req, res) => {
  try {
    const { fileName, holidayMap } = req.body as { fileName: string; holidayMap: Record<string, string[]> };
    const keys = Object.keys(holidayMap ?? {});
    const months = [...new Set(keys.map(k => k.slice(0, 2)))].sort();
    console.log(`[HolidayMap POST] 파일명: ${fileName}`);
    console.log(`[HolidayMap POST] 날짜 수: ${keys.length}, 월 목록: ${months.join(", ")}`);
    const now = new Date();
    await db
      .insert(holidayStoreTable)
      .values({ id: 1, fileName, holidayMap })
      .onConflictDoUpdate({
        target: holidayStoreTable.id,
        set: { fileName, holidayMap, updatedAt: now },
      });
    console.log(`[HolidayMap POST] DB 저장 완료: updatedAt=${now.toISOString()}`);
    res.json({ ok: true, months, keyCount: keys.length, updatedAt: now.toISOString() });
  } catch (err) {
    console.error(`[HolidayMap POST] 오류:`, err);
    res.status(500).json({ error: String(err) });
  }
});

export default router;
