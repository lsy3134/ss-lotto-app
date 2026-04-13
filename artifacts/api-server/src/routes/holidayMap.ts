import { Router, type IRouter } from "express";
import { db, holidayStoreTable } from "@workspace/db";
import { eq } from "drizzle-orm";

const router: IRouter = Router();

router.get("/holiday-map", async (_req, res) => {
  try {
    const rows = await db.select().from(holidayStoreTable).where(eq(holidayStoreTable.id, 1));
    if (rows.length === 0) {
      res.json({ fileName: "", holidayMap: {} });
      return;
    }
    res.json({ fileName: rows[0].fileName, holidayMap: rows[0].holidayMap });
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

router.post("/holiday-map", async (req, res) => {
  try {
    const { fileName, holidayMap } = req.body as { fileName: string; holidayMap: Record<string, string[]> };
    await db
      .insert(holidayStoreTable)
      .values({ id: 1, fileName, holidayMap })
      .onConflictDoUpdate({
        target: holidayStoreTable.id,
        set: { fileName, holidayMap, updatedAt: new Date() },
      });
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

export default router;
