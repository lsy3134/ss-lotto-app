import { Router, type IRouter } from "express";
import { pool } from "@workspace/db";

const router: IRouter = Router();

router.get("/device-access", async (_req, res) => {
  res.set("Cache-Control", "no-store, no-cache, must-revalidate");
  try {
    const [settings, bindings] = await Promise.all([
      pool.query<{ enabled: boolean }>(
        "SELECT enabled FROM device_access_settings WHERE id = 1",
      ),
      pool.query<{ userName: string; registeredAt: Date }>(`
        SELECT user_name AS "userName", registered_at AS "registeredAt"
        FROM user_device_bindings
        ORDER BY user_name
      `),
    ]);
    res.json({
      enabled: settings.rows[0]?.enabled ?? false,
      bindings: bindings.rows.map((row) => ({
        userName: row.userName,
        registeredAt: row.registeredAt.toISOString(),
      })),
    });
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

router.post("/device-access/check", async (req, res) => {
  const userName = typeof req.body?.userName === "string" ? req.body.userName.trim() : "";
  const deviceId = typeof req.body?.deviceId === "string" ? req.body.deviceId.trim() : "";
  if (!userName || !deviceId || deviceId.length > 200) {
    res.status(400).json({ allowed: false, reason: "잘못된 로그인 요청입니다." });
    return;
  }

  try {
    const settings = await pool.query<{ enabled: boolean }>(
      "SELECT enabled FROM device_access_settings WHERE id = 1",
    );
    if (!(settings.rows[0]?.enabled ?? false)) {
      res.json({ allowed: true, enabled: false, registered: false });
      return;
    }

    await pool.query(
      `INSERT INTO user_device_bindings (user_name, device_id)
       VALUES ($1, $2)
       ON CONFLICT (user_name) DO NOTHING`,
      [userName, deviceId],
    );
    const binding = await pool.query<{ deviceId: string }>(
      `SELECT device_id AS "deviceId" FROM user_device_bindings WHERE user_name = $1`,
      [userName],
    );
    const allowed = binding.rows[0]?.deviceId === deviceId;
    res.status(allowed ? 200 : 409).json({
      allowed,
      enabled: true,
      registered: allowed,
      reason: allowed ? undefined : "이 이름은 이미 다른 기기에 등록되어 있습니다. 관리자에게 기기 연결 해제를 요청해 주세요.",
    });
  } catch (err) {
    res.status(500).json({ allowed: false, reason: "기기 등록 확인 중 오류가 발생했습니다.", error: String(err) });
  }
});

router.post("/device-access/settings", async (req, res) => {
  if (typeof req.body?.enabled !== "boolean") {
    res.status(400).json({ error: "enabled 값이 필요합니다." });
    return;
  }
  try {
    await pool.query(
      `INSERT INTO device_access_settings (id, enabled, updated_at)
       VALUES (1, $1, now())
       ON CONFLICT (id) DO UPDATE SET enabled = EXCLUDED.enabled, updated_at = now()`,
      [req.body.enabled],
    );
    res.json({ ok: true, enabled: req.body.enabled });
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

router.delete("/device-access/:userName", async (req, res) => {
  try {
    await pool.query("DELETE FROM user_device_bindings WHERE user_name = $1", [req.params.userName]);
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

export default router;
