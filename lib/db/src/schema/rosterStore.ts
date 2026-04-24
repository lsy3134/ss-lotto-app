import { pgTable, serial, jsonb, timestamp } from "drizzle-orm/pg-core";

export const rosterStoreTable = pgTable("roster_store", {
  id: serial("id").primaryKey(),
  roster: jsonb("roster").notNull().default([]),
  updatedAt: timestamp("updated_at", { withTimezone: true }).defaultNow(),
});

export type RosterStore = typeof rosterStoreTable.$inferSelect;
