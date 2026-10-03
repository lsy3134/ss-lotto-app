import app from "./app";
import { logger } from "./lib/logger";
import { pool } from "@workspace/db";

const rawPort = process.env["PORT"];

if (!rawPort) {
  throw new Error(
    "PORT environment variable is required but was not provided.",
  );
}

const port = Number(rawPort);

if (Number.isNaN(port) || port <= 0) {
  throw new Error(`Invalid PORT value: "${rawPort}"`);
}

// 운영 DB에 roster_store 테이블이 없을 경우 자동 생성
async function ensureTables() {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS roster_store (
        id serial PRIMARY KEY,
        roster jsonb NOT NULL DEFAULT '[]',
        updated_at timestamptz DEFAULT now()
      );
      CREATE TABLE IF NOT EXISTS device_access_settings (
        id integer PRIMARY KEY,
        enabled boolean NOT NULL DEFAULT false,
        updated_at timestamptz NOT NULL DEFAULT now()
      );
      INSERT INTO device_access_settings (id, enabled)
      VALUES (1, false)
      ON CONFLICT (id) DO NOTHING;
      CREATE TABLE IF NOT EXISTS user_device_bindings (
        user_name text PRIMARY KEY,
        device_id text NOT NULL,
        registered_at timestamptz NOT NULL DEFAULT now()
      );
    `);
    logger.info("DB 테이블 확인 완료 (roster_store, device_access_settings, user_device_bindings)");
  } catch (err) {
    logger.error({ err }, "DB 테이블 생성 오류");
  }
}

ensureTables().then(() => {
  app.listen(port, (err) => {
    if (err) {
      logger.error({ err }, "Error listening on port");
      process.exit(1);
    }
    logger.info({ port }, "Server listening");
  });
});
