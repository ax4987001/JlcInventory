using System;
using System.IO;
using System.Data;
using ExcelDataReader;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using MiniExcelLibs;

namespace JlcInventory
{
    public static class ExcelHelper
    {
        private static DataTable ReadExcel(string path)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var result = reader.AsDataSet();
            return result.Tables[0];
        }

        private static int FindHeaderRow(DataTable table)
        {
            for (int i = 0; i < Math.Min(20, table.Rows.Count); i++)
            {
                int matchCount = 0;
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    string val = table.Rows[i][j]?.ToString()?.Trim();
                    if (val == "Supplier Part" || val == "商品编号" || val == "Designator" || val == "位号" || val == "Footprint" || val == "封装" || val == "Name" || val == "型号" || val == "Comment" || val == "Value") 
                    {
                        matchCount++;
                    }
                }
                if (matchCount >= 2) return i; // 至少有两个常见表头才认为是表头行
            }
            return -1;
        }

        private static Dictionary<string, int> GetColMap(DataTable table, int headerRow)
        {
            var map = new Dictionary<string, int>();
            for (int j = 0; j < table.Columns.Count; j++)
            {
                string val = table.Rows[headerRow][j]?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(val) && !map.ContainsKey(val)) map[val] = j;
            }
            return map;
        }

        private static string GetVal(DataRow row, Dictionary<string, int> map, params string[] keys)
        {
            foreach(var k in keys) {
                if (map.TryGetValue(k, out int idx)) {
                    var val = row[idx]?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(val)) return val;
                }
            }
            return "";
        }

        private static string NormalizeFootprint(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return "";
            raw = raw.ToUpper().Trim();

            // 1. 提取常见的贴片尺寸 (如 C0805, R0805, LED0805 -> 0805)
            var match = Regex.Match(raw, @"(0201|0402|0603|0805|1206|1210|2010|2512)");
            if (match.Success) return match.Value;

            // 2. 提取 SOT 系列 (如 SOT-23_... -> SOT-23)
            match = Regex.Match(raw, @"(SOT)-?(\d+)");
            if (match.Success) return $"{match.Groups[1].Value}-{match.Groups[2].Value}";

            // 3. 提取 SOP/SOIC 系列 (如 SOIC-16_... -> SOP-16)
            match = Regex.Match(raw, @"(SOIC|SOP|SSOP|TSSOP|MSOP)-?(\d+)");
            if (match.Success) {
                string type = match.Groups[1].Value;
                if (type == "SOIC") type = "SOP"; // 将 SOIC 归一化为 SOP
                return $"{type}-{match.Groups[2].Value}";
            }

            // 4. 提取 QFP/QFN 系列
            match = Regex.Match(raw, @"(LQFP|QFN|QFP|TQFP)-?(\d+)");
            if (match.Success) return $"{match.Groups[1].Value}-{match.Groups[2].Value}";

            // 5. 提取二极管等标准
            match = Regex.Match(raw, @"(SMA|SMB|SMC|DO-\d+\w*)");
            if (match.Success) return match.Value;

            // 6. 如果带有下划线 _ 且以上都不匹配，截取下划线前的核心部分
            int underscoreIdx = raw.IndexOf('_');
            if (underscoreIdx > 0) {
                return raw.Substring(0, underscoreIdx);
            }

            return raw;
        }

        private static string GuessCategoryFromFootprint(string footprint, string name)
        {
            if (string.IsNullOrEmpty(footprint) && string.IsNullOrEmpty(name)) return "";
            footprint = footprint.ToUpper().Trim();
            name = name.ToUpper().Trim();

            // 1. 基于嘉立创官方封装前缀的强匹配
            if (footprint.StartsWith("RES-") || Regex.IsMatch(footprint, @"^R\d{4}")) return "电阻";
            if (footprint.StartsWith("CAP-") || Regex.IsMatch(footprint, @"^C\d{4}")) return "电容";
            if (footprint.StartsWith("IND-") || Regex.IsMatch(footprint, @"^L\d{4}")) return "电感";
            if (footprint.StartsWith("BEAD-")) return "磁珠";
            if (footprint.StartsWith("FILTER-")) return "滤波器";
            if (footprint.StartsWith("FUSE-") || Regex.IsMatch(footprint, @"^F\d{4}")) return "保险丝";
            if (footprint.StartsWith("LED") || footprint.Contains("LED")) return "光电器件";
            if (Regex.IsMatch(footprint, @"^X\d{4}") || footprint.Contains("XTAL") || footprint.Contains("OSC")) return "振荡器";
            if (Regex.IsMatch(footprint, @"^D\d{4}") || footprint.StartsWith("DIODE") || footprint.StartsWith("SMA") || footprint.StartsWith("SMB") || footprint.StartsWith("SMC") || footprint.StartsWith("SOD-") || footprint.StartsWith("DO-")) {
                if (name.Contains("TVS")) return "TVS";
                return "二极管";
            }
            if (footprint.StartsWith("CONN") || footprint.StartsWith("HDR") || footprint.Contains("PIN") || footprint.Contains("XH") || footprint.Contains("PH") || footprint.Contains("CH") || footprint.Contains("KF")) return "连接器";
            if (footprint.Contains("RELAY")) return "继电器";
            if (footprint.Contains("SW") || footprint.Contains("KEY") || footprint.Contains("BUTTON")) return "开关";

            // 2. 基于型号 (Name) 的补充推测
            if (name.EndsWith("Ω") || name.EndsWith("OHM")) return "电阻";
            if (name.EndsWith("F") || name.EndsWith("FARAD")) return "电容";
            if (name.EndsWith("H") || name.EndsWith("HENRY")) return "电感";
            if (name.EndsWith("HZ")) return "振荡器";
            
            // 3. 泛半导体封装 (SOT, SOP, QFP, QFN, BGA, DIP 等) 的细化推断
            if (Regex.IsMatch(footprint, @"(SOT|SOP|SOIC|TSSOP|MSOP|SSOP|QFP|LQFP|TQFP|QFN|DFN|BGA|DIP|LCC|TO-\d+|IC|MCU)")) {
                if (name.Contains("MEM") || name.Contains("FLASH") || name.Contains("EEPROM") || name.StartsWith("AT24") || name.StartsWith("W25")) return "存储器";
                if (name.StartsWith("STM32") || name.StartsWith("ESP32") || name.StartsWith("PIC") || name.StartsWith("CH32") || name.StartsWith("ATMEGA") || name.StartsWith("GD32")) return "单片机";
                if (name.StartsWith("LM") || name.StartsWith("NE555") || name.StartsWith("OPA") || name.StartsWith("TL0")) return "运算放大器";
                if (name.StartsWith("MAX") || name.StartsWith("TPS") || name.StartsWith("LDO") || name.StartsWith("AMS1117") || name.StartsWith("LM78") || name.Contains("REGULATOR")) return "电源管理";
                if (name.StartsWith("CH340") || name.StartsWith("CP210") || name.StartsWith("MAX232") || name.StartsWith("SP3485")) return "通信接口芯片";
                if (name.StartsWith("TM16") || name.StartsWith("MAX7219")) return "数码管驱动";
                if (name.Contains("TRANSISTOR") || name.StartsWith("2N") || name.StartsWith("S9") || name.StartsWith("AO") || name.StartsWith("IRF") || name.StartsWith("BSS")) return "MOS管"; // 归入 MOS管 或 三极管，前端允许 "三极管" 或 "MOS管"
                if (name.Contains("ADC") || name.Contains("DAC")) return "ADC";
                if (name.Contains("RF") || name.Contains("NRF") || name.Contains("CC25")) return "射频芯片";
                
                // 默认退化到逻辑器件
                return "逻辑器件";
            }
            
            // 4. 其他后备策略
            if (footprint.Contains("RES") || footprint.Contains("R0") || footprint.Contains("R1") || footprint.Contains("R2") || name.EndsWith("R")) return "电阻";
            if (footprint.Contains("CAP") || footprint.Contains("C0") || footprint.Contains("C1") || footprint.Contains("C2")) return "电容";
            if (footprint.Contains("IND") || footprint.Contains("L0") || footprint.Contains("L1") || footprint.Contains("L2")) return "电感";

            return "";
        }

        public static object ParseInventoryExcel(string path)
        {
            try {
                var table = ReadExcel(path);
                int headerRow = FindHeaderRow(table);
                // 移除强校验，允许无表头的尝试解析或者将第一行视为表头
                if (headerRow == -1) headerRow = 0; // 如果实在找不到，强行使用第一行

                var map = GetColMap(table, headerRow);
                var results = new List<object>();
                for (int i = headerRow + 1; i < table.Rows.Count; i++)
                {
                    var row = table.Rows[i];
                    string category = GetVal(row, map, "目录", "Category");
                    string name = GetVal(row, map, "Name", "型号", "Comment", "Value");
                    string footprint = NormalizeFootprint(GetVal(row, map, "Footprint", "封装"));
                    
                    if (string.IsNullOrEmpty(category)) category = GuessCategoryFromFootprint(GetVal(row, map, "Footprint", "封装"), name);

                    string qtyStr = GetVal(row, map, "购买数量", "订购数量", "Quantity", "所需数量", "库存");
                    int qty = 0;
                    if (double.TryParse(qtyStr, out double dQty)) qty = (int)dQty;

                    string designator = GetVal(row, map, "Designator", "位号");
                    if (!string.IsNullOrEmpty(designator)) {
                        int trueReqQty = designator.Split(new[] { ',', '，', ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;
                        if (qty == 0) qty = trueReqQty;
                    }

                    string priceStr = GetVal(row, map, "单价(RMB)", "Price");
                    double price = 0;
                    double.TryParse(priceStr, out price);

                    if (qty > 0 && (!string.IsNullOrEmpty(name) || !string.IsNullOrEmpty(footprint)))
                    {
                        if (string.IsNullOrEmpty(name)) name = footprint; // Fallback
                        results.Add(new {
                            jlc_id = "",
                            category = category,
                            name = name,
                            footprint = footprint,
                            quantity = qty,
                            price = price
                        });
                    }
                }
                
                if (results.Count == 0) throw new Exception("未能从表格中解析出任何有效元器件数据，请检查表头是否包含“型号/Name”、“封装/Footprint”及“数量/Quantity”。");
                
                return new { status = "success", data = results };
            } catch (Exception ex) {
                return new { status = "error", message = ex.Message };
            }
        }

        public static object CompareBom(string path)
        {
            try {
                var table = ReadExcel(path);
                int headerRow = FindHeaderRow(table);
                if (headerRow == -1) throw new Exception("无法识别表头。");

                var map = GetColMap(table, headerRow);
                var allInv = DbHelper.GetAll();
                
                var invByJlc = new Dictionary<string, InventoryItem>();
                var invByNameFp = new Dictionary<string, InventoryItem>();
                var invByNameNormFp = new Dictionary<string, InventoryItem>();

                foreach(var item in allInv) {
                    if (!string.IsNullOrEmpty(item.jlc_id)) invByJlc[item.jlc_id.ToUpper()] = item;
                    if (!string.IsNullOrEmpty(item.name)) {
                        invByNameFp[$"{item.name.ToUpper()}_{item.footprint.ToUpper()}"] = item;
                        invByNameNormFp[$"{item.name.ToUpper()}_{NormalizeFootprint(item.footprint)}"] = item;
                    }
                }

                var results = new List<object>();
                for (int i = headerRow + 1; i < table.Rows.Count; i++)
                {
                    var row = table.Rows[i];
                    string jlc = GetVal(row, map, "商品编号", "Supplier Part");
                    string category = GetVal(row, map, "目录", "Category");
                    string name = GetVal(row, map, "Name", "型号", "Comment", "Value");
                    string footprint = NormalizeFootprint(GetVal(row, map, "Footprint", "封装"));
                    
                    if (string.IsNullOrEmpty(category)) category = GuessCategoryFromFootprint(GetVal(row, map, "Footprint", "封装"), name);
                    if (string.IsNullOrEmpty(name)) name = footprint;
                    
                    string designator = GetVal(row, map, "Designator", "位号");
                    int trueReqQty = 0;
                    if (!string.IsNullOrEmpty(designator)) {
                        trueReqQty = designator.Split(new[] { ',', '，', ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;
                    }
                    
                    string reqQtyStr = GetVal(row, map, "所需数量", "数量", "Quantity");
                    if (trueReqQty == 0 && double.TryParse(reqQtyStr, out double dR)) trueReqQty = (int)dR;

                    string moqQtyStr = GetVal(row, map, "购买数量", "订购数量");
                    int moqQty = 0;
                    if (double.TryParse(moqQtyStr, out double dMoq)) moqQty = (int)dMoq;
                    if (moqQty == 0) moqQty = trueReqQty; // 降级为真实需求量
                    if (trueReqQty == 0) trueReqQty = moqQty; // 互相兜底
                    
                    if (trueReqQty == 0 && moqQty == 0) continue;

                    string priceStr = GetVal(row, map, "单价(RMB)", "Price");
                    double jlcUnitPrice = 0;
                    double.TryParse(priceStr, out jlcUnitPrice);

                    string subTotalStr = GetVal(row, map, "小计(RMB)", "小计");
                    double jlcSubTotal = 0;
                    if (double.TryParse(subTotalStr, out double dSub)) {
                        jlcSubTotal = dSub;
                        if (jlcUnitPrice == 0 && moqQty > 0) jlcUnitPrice = jlcSubTotal / moqQty; // 倒推单价
                    } else {
                        jlcSubTotal = jlcUnitPrice * moqQty;
                    }

                    InventoryItem localItem = null;
                    if (!string.IsNullOrEmpty(jlc) && invByJlc.TryGetValue(jlc.ToUpper(), out var byJlc)) {
                        localItem = byJlc;
                    } else {
                        string keyExact = $"{name.ToUpper()}_{footprint.ToUpper()}";
                        string keyNorm = $"{name.ToUpper()}_{NormalizeFootprint(footprint)}";

                        if (invByNameFp.TryGetValue(keyExact, out var byName)) {
                            localItem = byName;
                        } else if (invByNameNormFp.TryGetValue(keyNorm, out var byNorm)) {
                            localItem = byNorm;
                        }
                    }

                    string status = "red";
                    int localQty = 0;
                    double localPrice = 0;

                    if (localItem != null) {
                        localQty = localItem.quantity;
                        localPrice = localItem.price;

                        if (localQty >= trueReqQty) {
                            if (jlcUnitPrice == 0 && jlcSubTotal == 0) {
                                status = "green";
                            } else {
                                // 本地总价 = 真实需求量 * 本地单价
                                // JLC 总价 = jlcSubTotal
                                if ((localPrice * trueReqQty) <= jlcSubTotal || localPrice == 0) status = "green";
                                else status = "yellow";
                            }
                        } else {
                            status = "red";
                        }
                    }

                    results.Add(new {
                        jlc_id = jlc,
                        category = category,
                        name = name,
                        footprint = footprint,
                        req_qty = trueReqQty,
                        moq_qty = moqQty,
                        jlc_price = jlcUnitPrice,
                        jlc_subtotal = jlcSubTotal,
                        local_qty = localQty,
                        local_price = localPrice,
                        status = status,
                        local_item_id = localItem != null ? localItem.id : 0
                    });
                }

                return new { status = "success", data = results };
            } catch (Exception ex) {
                return new { status = "error", message = ex.Message };
            }
        }

        public static object ExportExcel(string savePath)
        {
            try {
                var items = DbHelper.GetAll();
                var exportData = new List<object>();
                foreach(var item in items) {
                    exportData.Add(new {
                        型号 = item.name,
                        封装 = item.footprint,
                        目录 = item.category,
                        库存 = item.quantity,
                        单价 = item.price
                    });
                }
                MiniExcel.SaveAs(savePath, exportData);
                return new { status = "success", message = "导出成功！" };
            } catch(Exception ex) {
                return new { status = "error", message = ex.Message };
            }
        }
    }
}