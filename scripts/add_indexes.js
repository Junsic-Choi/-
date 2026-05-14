
require('dotenv').config();
const { createClient } = require('@libsql/client');

const client = createClient({
    url: process.env.DATABASE_URL,
    authToken: process.env.DATABASE_AUTH_TOKEN,
});

async function addIndexes() {
    console.log("Adding indexes to 'plans' table...");
    try {
        // Index for filtering by equipment and week (used in most queries)
        await client.execute("CREATE INDEX IF NOT EXISTS idx_plans_eq_week ON plans(equipment, weekId)");
        console.log("Created index: idx_plans_eq_week");

        // Index for weekId (used in subqueries and consolidated view)
        await client.execute("CREATE INDEX IF NOT EXISTS idx_plans_weekId ON plans(weekId)");
        console.log("Created index: idx_plans_weekId");
        
        // Index for manager (used for filtering)
        await client.execute("CREATE INDEX IF NOT EXISTS idx_plans_manager ON plans(manager)");
        console.log("Created index: idx_plans_manager");

        console.log("All indexes added successfully.");
    } catch (err) {
        console.error("Error adding indexes:", err);
    }
}

addIndexes();
