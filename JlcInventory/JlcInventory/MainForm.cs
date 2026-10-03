using System;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using System.Data;

namespace JlcInventory
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            this.Text = "嘉立创库存计价器";
            this.Width = 1050;
            this.Height = 700;
            this.StartPosition = FormStartPosition.CenterScreen;
            DbHelper.InitDb(); // 初始化数据库
            InitializeWebView();
        }

        private async void InitializeWebView()
        {
            var webView = new Microsoft.Web.WebView2.WinForms.WebView2();
            webView.Dock = DockStyle.Fill;
            this.Controls.Add(webView);
            
            var userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "JlcInventoryWebView2");
            var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
            await webView.EnsureCoreWebView2Async(env);

            webView.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
            
            string htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "index.html");
            webView.CoreWebView2.Navigate(htmlPath);
        }

        private void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            var msg = e.TryGetWebMessageAsString();
            var req = JsonSerializer.Deserialize<ApiRequest>(msg);
            object resultData = null;

            try
            {
                if (req.method == "SelectFile")
                {
                    using (var dlg = new OpenFileDialog { Filter = "Excel Files|*.xls;*.xlsx|All files|*.*" })
                    {
                        resultData = dlg.ShowDialog() == DialogResult.OK ? dlg.FileName : null;
                    }
                }
                else if (req.method == "GetInventory")
                {
                    resultData = DbHelper.GetAll();
                }
                else if (req.method == "GetLogs")
                {
                    resultData = DbHelper.GetLogs();
                }
                else if (req.method == "RollbackLog")
                {
                    var id = req.args.Deserialize<long>();
                    DbHelper.RollbackLog(id);
                    resultData = new { status = "success" };
                }
                else if (req.method == "AddSingle")
                {
                    var args = req.args.Deserialize<AddSingleArgs>();
                    string diff = DbHelper.AddOrUpdate(args.jlc_id, args.category, args.name, args.footprint, args.quantity, args.price);
                    DbHelper.AddLog("添加单项", $"添加了 {args.name} ({args.footprint}) x{args.quantity}", diff);
                    resultData = new { status = "success" };
                }
                else if (req.method == "BatchAdd")
                {
                    var payload = req.args.Deserialize<BatchAddPayload>();
                    List<string> diffs = new List<string>();
                    foreach(var item in payload.items) {
                        diffs.Add(DbHelper.AddOrUpdate(item.jlc_id, item.category, item.name, item.footprint, item.quantity, item.price));
                    }
                    string src = string.IsNullOrEmpty(payload.filename) ? "文本批量录入" : $"文件导入: {Path.GetFileName(payload.filename)}";
                    string descCats = payload.items.Count > 0 ? $"包含了 {payload.items[0].category} 等" : "";
                    DbHelper.AddLog("批量添加", $"来源: {src}，共 {payload.items.Count} 项，{descCats}", string.Join("\n", diffs));
                    resultData = new { status = "success" };
                }
                else if (req.method == "ImportExcel")
                {
                    var file = req.args.Deserialize<string>();
                    resultData = ExcelHelper.ParseInventoryExcel(file);
                }
                else if (req.method == "CompareBom")
                {
                    var file = req.args.Deserialize<string>();
                    resultData = ExcelHelper.CompareBom(file);
                }
                else if (req.method == "UpdateItem")
                {
                    var args = req.args.Deserialize<UpdateItemArgs>();
                    string diff = DbHelper.UpdateItemField(args.id, args.field, args.value);
                    DbHelper.AddLog("修改单项", $"修改了 ID {args.id} 的 {args.field} 为 {args.value}", diff);
                    resultData = new { status = "success" };
                }
                else if (req.method == "ExportExcel")
                {
                    using (var dlg = new SaveFileDialog { Filter = "Excel Files|*.xlsx", FileName = "本地库存导出.xlsx" })
                    {
                        if (dlg.ShowDialog() == DialogResult.OK)
                        {
                            resultData = ExcelHelper.ExportExcel(dlg.FileName);
                        }
                    }
                }
                else if (req.method == "SettleInventory")
                {
                    var payload = req.args.Deserialize<SettlePayload>();
                    List<string> diffs = new List<string>();
                    foreach (var item in payload.items)
                    {
                        diffs.Add(DbHelper.DeductInventory(item.id, item.deduct_qty));
                    }
                    DbHelper.AddLog("一键结算", $"扣除了 {payload.items.Count} 项库存", $"基于比对表: {Path.GetFileName(payload.filename)}\n\n{string.Join("\n", diffs)}");
                    resultData = new { status = "success" };
                }
            }
            catch (Exception ex)
            {
                resultData = new { status = "error", message = ex.Message };
            }

            var response = new { cbId = req.cbId, data = resultData };
            var responseJson = JsonSerializer.Serialize(response);
            ((CoreWebView2)sender).PostWebMessageAsString(responseJson);
        }
    }

    public class ApiRequest
    {
        public string method { get; set; }
        public JsonElement args { get; set; }
        public int cbId { get; set; }
    }

    public class AddSingleArgs
    {
        public string jlc_id { get; set; }
        public string category { get; set; }
        public string name { get; set; }
        public string footprint { get; set; }
        public int quantity { get; set; }
        public double price { get; set; }
    }

    public class UpdateItemArgs
    {
        public long id { get; set; }
        public string field { get; set; }
        public string value { get; set; }
    }

    public class BatchAddPayload
    {
        public List<AddSingleArgs> items { get; set; }
        public string filename { get; set; }
    }

    public class SettleItem
    {
        public long id { get; set; }
        public int deduct_qty { get; set; }
    }

    public class SettlePayload
    {
        public string filename { get; set; }
        public List<SettleItem> items { get; set; }
    }
}