using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.Data.Sqlite;
using System.Text.Json;

namespace JlcInventory
{
    public class InventoryItem
    {
        public long id { get; set; }
        public string jlc_id { get; set; }
        public string category { get; set; }
        public string name { get; set; }
        public string footprint { get; set; }
        public int quantity { get; set; }
        public double price { get; set; }
    }

    public class LogItem
    {
        public long id { get; set; }
        public string timestamp { get; set; }
        public string action_type { get; set; }
        public string description { get; set; }
        public string details { get; set; }
        public string snapshot { get; set; }
    }

    public static class DbHelper
    {
        static string DbPath {
            get {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string configPath = Path.Combine(baseDir, "config.json");
                if (File.Exists(configPath)) {
                    try {
                        var json = File.ReadAllText(configPath);
                        using (var doc = JsonDocument.Parse(json)) {
                            if (doc.RootElement.TryGetProperty("DataPath", out var dataPathProp)) {
                                string dataPath = dataPathProp.GetString();
                                if (!string.IsNullOrEmpty(dataPath)) {
                                    if (!Directory.Exists(dataPath)) Directory.CreateDirectory(dataPath);
                                    return Path.Combine(dataPath, "inventory.db");
                                }
                            }
                        }
                    } catch { }
                }
                return Path.Combine(baseDir, "inventory.db");
            }
        }
        static string connStr => $"Data Source={DbPath}";
        
        public static void InitDb()
        {
            using var conn = new SqliteConnection(connStr);
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                CREATE TABLE IF NOT EXISTS inventory (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    jlc_id TEXT,
                    category TEXT,
                    name TEXT,
                    footprint TEXT,
                    quantity INTEGER DEFAULT 0,
                    price REAL DEFAULT 0.0
                );
                CREATE INDEX IF NOT EXISTS idx_jlc_id ON inventory(jlc_id);
                CREATE INDEX IF NOT EXISTS idx_name_footprint ON inventory(name, footprint);

                CREATE TABLE IF NOT EXISTS logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                    action_type TEXT,
                    description TEXT,
                    details TEXT,
                    snapshot TEXT
                );
            ";
            cmd.ExecuteNonQuery();

            try {
                using var cmdAlter = conn.CreateCommand();
                cmdAlter.CommandText = "ALTER TABLE inventory ADD COLUMN category TEXT DEFAULT ''";
                cmdAlter.ExecuteNonQuery();
            } catch { }
        }

        public static string GetSnapshot()
        {
            var items = GetAll();
            return JsonSerializer.Serialize(items);
        }

        public static void AddLog(string actionType, string description, string details)
        {
            try {
                string snapshot = GetSnapshot();
                using var conn = new SqliteConnection(connStr);
                conn.Open();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "INSERT INTO logs (timestamp, action_type, description, details, snapshot) VALUES (datetime('now', 'localtime'), @at, @desc, @det, @snap)";
                cmd.Parameters.AddWithValue("@at", actionType);
                cmd.Parameters.AddWithValue("@desc", description);
                cmd.Parameters.AddWithValue("@det", details);
                cmd.Parameters.AddWithValue("@snap", snapshot);
                cmd.ExecuteNonQuery();
            } catch (Exception ex) {
                Console.WriteLine("AddLog Error: " + ex.Message);
            }
        }

        public static List<LogItem> GetLogs()
        {
            var list = new List<LogItem>();
            using var conn = new SqliteConnection(connStr);
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT id, timestamp, action_type, description, details, snapshot FROM logs ORDER BY id DESC";
            using var reader = cmd.ExecuteReader();
            while(reader.Read())
            {
                list.Add(new LogItem {
                    id = reader.GetInt64(0),
                    timestamp = reader.IsDBNull(1) ? "" : reader.GetString(1),
                    action_type = reader.IsDBNull(2) ? "" : reader.GetString(2),
                    description = reader.IsDBNull(3) ? "" : reader.GetString(3),
                    details = reader.IsDBNull(4) ? "" : reader.GetString(4),
                    snapshot = reader.IsDBNull(5) ? "" : reader.GetString(5)
                });
            }
            return list;
        }

        public static void RollbackLog(long logId)
        {
            using var conn = new SqliteConnection(connStr);
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT snapshot FROM logs WHERE id = @id";
            cmd.Parameters.AddWithValue("@id", logId);
            string snapshot = cmd.ExecuteScalar() as string;
            
            if (!string.IsNullOrEmpty(snapshot))
            {
                var items = JsonSerializer.Deserialize<List<InventoryItem>>(snapshot);
                using var tx = conn.BeginTransaction();
                try {
                    using var cmdClear = conn.CreateCommand();
                    cmdClear.CommandText = "DELETE FROM inventory";
                    cmdClear.ExecuteNonQuery();

                    using var cmdInsert = conn.CreateCommand();
                    cmdInsert.CommandText = "INSERT INTO inventory (id, jlc_id, category, name, footprint, quantity, price) VALUES (@id, @jlc, @c, @n, @f, @q, @p)";
                    foreach(var item in items) {
                        cmdInsert.Parameters.Clear();
                        cmdInsert.Parameters.AddWithValue("@id", item.id);
                        cmdInsert.Parameters.AddWithValue("@jlc", item.jlc_id ?? "");
                        cmdInsert.Parameters.AddWithValue("@c", item.category ?? "");
                        cmdInsert.Parameters.AddWithValue("@n", item.name ?? "");
                        cmdInsert.Parameters.AddWithValue("@f", item.footprint ?? "");
                        cmdInsert.Parameters.AddWithValue("@q", item.quantity);
                        cmdInsert.Parameters.AddWithValue("@p", item.price);
                        cmdInsert.ExecuteNonQuery();
                    }
                    tx.Commit();
                    AddLog("Rollback", $"回退至日志 #{logId} 的状态", "执行了快照回滚操作");
                } catch {
                    tx.Rollback();
                    throw;
                }
            }
        }

        public static List<InventoryItem> GetAll()
        {
            var list = new List<InventoryItem>();
            using var conn = new SqliteConnection(connStr);
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT id, jlc_id, category, name, footprint, quantity, price FROM inventory";
            using var reader = cmd.ExecuteReader();
            while(reader.Read())
            {
                list.Add(new InventoryItem {
                    id = reader.GetInt64(0),
                    jlc_id = reader.IsDBNull(1) ? "" : reader.GetString(1),
                    category = reader.IsDBNull(2) ? "" : reader.GetString(2),
                    name = reader.IsDBNull(3) ? "" : reader.GetString(3),
                    footprint = reader.IsDBNull(4) ? "" : reader.GetString(4),
                    quantity = reader.IsDBNull(5) ? 0 : reader.GetInt32(5),
                    price = reader.IsDBNull(6) ? 0.0 : reader.GetDouble(6)
                });
            }
            return list;
        }

        public static string AddOrUpdate(string jlc_id, string category, string name, string footprint, int quantity, double price)
        {
            using var conn = new SqliteConnection(connStr);
            conn.Open();
            using var cmd = conn.CreateCommand();
            
            jlc_id = jlc_id?.Trim() ?? "";
            category = category?.Trim() ?? "";
            name = name?.Trim() ?? "";
            footprint = footprint?.Trim() ?? "";

            if (!string.IsNullOrEmpty(jlc_id)) {
                cmd.CommandText = "SELECT id, quantity, price, category FROM inventory WHERE jlc_id = @jlc_id";
                cmd.Parameters.AddWithValue("@jlc_id", jlc_id);
            } else {
                cmd.CommandText = "SELECT id, quantity, price, category FROM inventory WHERE name = @name AND footprint = @footprint";
                cmd.Parameters.AddWithValue("@name", name);
                cmd.Parameters.AddWithValue("@footprint", footprint);
            }

            long existingId = -1;
            int extQty = 0;
            double extPrice = 0;
            string extCat = "";

            using (var reader = cmd.ExecuteReader()) {
                if (reader.Read()) {
                    existingId = reader.GetInt64(0);
                    extQty = reader.GetInt32(1);
                    extPrice = reader.GetDouble(2);
                    extCat = reader.IsDBNull(3) ? "" : reader.GetString(3);
                }
            }

            using var cmdUpdate = conn.CreateCommand();
            string logDetail = "";
            if (existingId != -1) {
                double newPrice = extPrice;
                if (price > 0) {
                    int validExtQty = Math.Max(0, extQty);
                    int validAddQty = Math.Max(0, quantity);
                    if (validExtQty + validAddQty > 0) {
                        newPrice = ((validExtQty * extPrice) + (validAddQty * price)) / (validExtQty + validAddQty);
                    } else {
                        newPrice = price;
                    }
                }

                cmdUpdate.CommandText = "UPDATE inventory SET quantity = @q, price = @p, category = @c WHERE id = @id";
                cmdUpdate.Parameters.AddWithValue("@q", extQty + quantity);
                cmdUpdate.Parameters.AddWithValue("@p", newPrice);
                cmdUpdate.Parameters.AddWithValue("@c", string.IsNullOrEmpty(category) ? extCat : category);
                cmdUpdate.Parameters.AddWithValue("@id", existingId);
                logDetail = $"~ {category} {name} ({footprint}) 数量: {extQty} -> {extQty + quantity}, 单价: {extPrice:F4} -> {newPrice:F4}";
            } else {
                cmdUpdate.CommandText = "INSERT INTO inventory (jlc_id, category, name, footprint, quantity, price) VALUES (@jlc, @c, @n, @f, @q, @p)";
                cmdUpdate.Parameters.AddWithValue("@jlc", jlc_id);
                cmdUpdate.Parameters.AddWithValue("@c", category);
                cmdUpdate.Parameters.AddWithValue("@n", name);
                cmdUpdate.Parameters.AddWithValue("@f", footprint);
                cmdUpdate.Parameters.AddWithValue("@q", quantity);
                cmdUpdate.Parameters.AddWithValue("@p", price);
                logDetail = $"+ {category} {name} ({footprint}) 数量: +{quantity}, 单价: {price:F4}";
            }
            cmdUpdate.ExecuteNonQuery();
            return logDetail;
        }

        public static string UpdateItemField(long id, string field, string value)
        {
            using var conn = new SqliteConnection(connStr);
            conn.Open();
            using var cmd = conn.CreateCommand();
            var allowedFields = new List<string> { "jlc_id", "category", "name", "footprint", "quantity", "price" };
            if (!allowedFields.Contains(field)) return "";
            
            cmd.CommandText = $"SELECT category, name, footprint, {field} FROM inventory WHERE id = @id";
            cmd.Parameters.AddWithValue("@id", id);
            string cat="", name="", fp="", oldVal="";
            using (var reader = cmd.ExecuteReader()) {
                if (reader.Read()) {
                    cat = reader.IsDBNull(0) ? "" : reader.GetString(0);
                    name = reader.IsDBNull(1) ? "" : reader.GetString(1);
                    fp = reader.IsDBNull(2) ? "" : reader.GetString(2);
                    oldVal = reader.GetValue(3).ToString();
                }
            }

            using var cmdUpdate = conn.CreateCommand();
            cmdUpdate.CommandText = $"UPDATE inventory SET {field} = @v WHERE id = @id";
            cmdUpdate.Parameters.AddWithValue("@v", value);
            cmdUpdate.Parameters.AddWithValue("@id", id);
            cmdUpdate.ExecuteNonQuery();

            return $"~ {cat} {name} ({fp}) {field}: {oldVal} -> {value}";
        }

        public static string DeductInventory(long id, int deductQty)
        {
            using var conn = new SqliteConnection(connStr);
            conn.Open();
            using var cmd = conn.CreateCommand();
            
            cmd.CommandText = "SELECT category, name, footprint, quantity FROM inventory WHERE id = @id";
            cmd.Parameters.AddWithValue("@id", id);
            string cat="", name="", fp="";
            int extQty = 0;
            using (var reader = cmd.ExecuteReader()) {
                if (reader.Read()) {
                    cat = reader.IsDBNull(0) ? "" : reader.GetString(0);
                    name = reader.IsDBNull(1) ? "" : reader.GetString(1);
                    fp = reader.IsDBNull(2) ? "" : reader.GetString(2);
                    extQty = reader.GetInt32(3);
                }
            }

            using var cmdUpdate = conn.CreateCommand();
            cmdUpdate.CommandText = "UPDATE inventory SET quantity = quantity - @q WHERE id = @id";
            cmdUpdate.Parameters.AddWithValue("@q", deductQty);
            cmdUpdate.Parameters.AddWithValue("@id", id);
            cmdUpdate.ExecuteNonQuery();

            return $"- {cat} {name} ({fp}) 数量: {extQty} -> {extQty - deductQty}";
        }
    }
}