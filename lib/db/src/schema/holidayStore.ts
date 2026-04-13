import { pgTable, serial, text, jsonb, timestamp } from "drizzle-orm/pg-core";

export const holidayStoreTable = pgTable("holiday_store", {
  id: serial("id").primaryKey(),
  fileName: text("file_name").notNull().default(""),
  holidayMap: jsonb("holiday_map").notNull().default({}),
  updatedAt: timestamp("updated_at", { withTimezone: true }).defaultNow(),
});

export type HolidayStore = typeof holidayStoreTable.$inferSelect;
