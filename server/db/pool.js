import { Pool } from "pg";
import { DATABASE_POOL_MAX, DATABASE_URL } from "../config.js";

export const pool = new Pool({
  connectionString: DATABASE_URL,
  max: DATABASE_POOL_MAX
});
