import { Router, type IRouter } from "express";
import { db, rosterStoreTable } from "@workspace/db";
import { eq } from "drizzle-orm";

const router: IRouter = Router();

router.get("/roster", async (_req, res) => {
  res.set("Cache-Control", "no-store, no-cache, must-revalidate");
  res.set("Pragma", "no-cache");
  res.set("Expires", "0");
  try {
    const rows = await db.select().from(rosterStoreTable).where(eq(rosterStoreTable.id, 1));
    if (rows.length === 0) {
      res.json({ roster: [], updatedAt: null });
      return;
    }
    res.json({
      roster: rows[0].roster,
      updatedAt: rows[0].updatedAt ? rows[0].updatedAt.toISOString() : null,
    });
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

router.post("/roster", async (req, res) => {
  try {
    const { roster } = req.body as { roster: unknown[] };
    const now = new Date();
    console.log(`[Roster POST] 인원 수: ${roster?.length ?? 0}`);
    await db
      .insert(rosterStoreTable)
      .values({ id: 1, roster })
      .onConflictDoUpdate({
        target: rosterStoreTable.id,
        set: { roster, updatedAt: now },
      });
    console.log(`[Roster POST] DB 저장 완료: updatedAt=${now.toISOString()}`);
    res.json({ ok: true, count: roster?.length ?? 0, updatedAt: now.toISOString() });
  } catch (err) {
    console.error(`[Roster POST] 오류:`, err);
    res.status(500).json({ error: String(err) });
  }
});

export default router;
