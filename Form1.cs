using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
//using Outlook = Microsoft.Office.Interop.Outlook;
using ExcelDataReader;
using ClosedXML.Excel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

//using LicenseContext = System.ComponentModel.LicenseContext;

namespace ScheduleViewer
{
    public partial class Schedule_Viewer : Form
    {
        private string excelFilePath = @"K:\NewTestRequests\Default\Ply_Dyno_Schedule.xlsx"; // Set your Excel path
        private Timer refreshTimer;
        private DataTable detailsTable = new DataTable();
        private DateTime lastSentTime = DateTime.MinValue;
        private string lastUpdateFile = "lastUpdate.txt";
        private bool _isDriveDisconnected = false;
        //private Label lblLastUpdated;
        //private FileSystemWatcher fsWatcher;

        private readonly int[] scheduledHours = { 9, 15 }; // 9 AM, 3 PM
        private bool emailSentThisSlot = false; // prevent duplicate emails in the same slot
        private int dgvSummaryOriginalTop;
        private int dgvSummaryOriginalHeight;
        private Panel pnlHoldBanner;
        private Label lblHoldBanner;
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private const int WM_SETREDRAW = 0x000B;

        private void BeginGridUpdate(DataGridView dgv)
        {
            if (!dgv.IsHandleCreated) dgv.CreateControl();
            SendMessage(dgv.Handle, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
            dgv.SuspendLayout();
        }

        private void EndGridUpdate(DataGridView dgv)
        {
            dgv.ResumeLayout();
            SendMessage(dgv.Handle, WM_SETREDRAW, new IntPtr(1), IntPtr.Zero);
            dgv.Refresh(); // single repaint (fast)
        }

        // Parses "Yes - Repair" / "Yes-Engineering" -> "Repair" / "Engineering"
        private string ExtractHoldReason(string holdVal)
        {
            if (string.IsNullOrWhiteSpace(holdVal)) return "";

            string s = holdVal.Trim();
            if (!s.StartsWith("Yes", StringComparison.OrdinalIgnoreCase)) return "";

            int dash = s.IndexOf('-');
            if (dash >= 0 && dash < s.Length - 1)
                return s.Substring(dash + 1).Trim();

            // fallback if user typed only "Yes Repair" etc.
            return s.Substring(3).Trim();
        }
        public Schedule_Viewer()
        {
            InitializeComponent();
            //System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            if (File.Exists(lastUpdateFile))
            {
                DateTime.TryParse(File.ReadAllText(lastUpdateFile), out lastSentTime);
            }
            //refreshTimer = new Timer();
            //refreshTimer.Interval = 30000; // 0.5 min
            //refreshTimer.Tick += RefreshTimer_Tick;
            //refreshTimer.Start();
        }
        

        private void Schedule_Viewer_Load(object sender, EventArgs e)
        {

            this.WindowState = FormWindowState.Maximized;
            this.AutoScaleMode = AutoScaleMode.Dpi;

            // Set fonts ONCE (prevents tiny-font issue)
            dgvSummary.Font = new Font("Segoe UI", 10f, FontStyle.Regular);
            dgvSummary.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10f, FontStyle.Bold);

            dgvDetails.Font = new Font("Segoe UI", 10f, FontStyle.Regular);
            dgvDetails.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10f, FontStyle.Bold);

            // Avoid expensive autosizing that causes flicker
            dgvSummary.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dgvDetails.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvSummary.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dgvDetails.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;

            //if (dgvSummary.Columns.Contains("Test Script"))
            //    dgvSummary.Columns["Test Script"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //if (dgvSummary.Columns.Contains("Hold Dyno For Engineering/ Repair/ Maintenance"))
            //{
            //    var c = dgvSummary.Columns["Hold Dyno For Engineering/ Repair/ Maintenance"];
            //    c.HeaderText = "Dyno Hold";
            //    c.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            //    c.Width = 150;
            //    c.HeaderCell.Style.WrapMode = DataGridViewTriState.True;
            //}
            // Paint-based HOLD banners (only once)
            dgvSummary.Paint -= DgvSummary_PaintHoldBanners;
            dgvSummary.Paint += DgvSummary_PaintHoldBanners;

            // Resize: only resize columns (NOT rows)
            this.Resize += (s, ev) =>
            {
                dgvSummary.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dgvDetails.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            };

            LoadSummary();
            LoadDynoList();

            detailsTable = CreateDataTable();
            dgvDetails.DataSource = detailsTable;

            lastTop4RowsHash = GetTop4RowsHash();

            SetupTimer();

        }
        private void SetupTimer()
        {
            if (refreshTimer != null) return;

            refreshTimer = new Timer();
            refreshTimer.Interval = 30000;
            refreshTimer.Tick += RefreshTimer_Tick;
            refreshTimer.Start();
        }

        private void RefreshTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                DateTime wt = File.GetLastWriteTime(excelFilePath);

                // Auto-recovery: if we were disconnected, clear the indicator
                if (_isDriveDisconnected)
                {
                    _isDriveDisconnected = false;
                }

                if (wt == lastSeenWriteTime) return;
                lastSeenWriteTime = wt;
                LoadSummary();
                if (cmbDyno.SelectedItem != null)
                {
                    LoadDetails(cmbDyno.SelectedItem.ToString(), preserveSearch: true);
                }
                this.Text = $"Schedule Viewer v1.0 - Last Updated: {wt}";
                CheckAndSendEmailIfUpdated();
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                _isDriveDisconnected = true;
                this.Text = "Schedule Viewer v1.0 - Drive disconnected - retrying...";
                // Grid data is untouched — last loaded data remains visible
            }
        }

        private void CheckAndSendScheduledEmail()
        {
            DateTime now = DateTime.Now;
            if (scheduledHours.Contains(now.Hour) && now.Minute == 0)
            {
                if (!emailSentThisSlot) // send only once in this slot
                {
                    CheckAndSendEmailIfUpdated();
                    emailSentThisSlot = true;
                }
            }
            else
            {
                // Reset once we are out of the scheduled hour
                emailSentThisSlot = false;
            }
        }

        private DateTime lastFileModifiedTime = DateTime.MinValue;
        private readonly TimeSpan stabilityDelay = TimeSpan.FromMinutes(2); // wait 2 minutes after last change
        private bool waitingForClose = false;
        private string lastTop4RowsHash = "";
        

        private void CheckAndSendEmailIfUpdated()
        {

            
            if (!File.Exists(excelFilePath)) return;

            if (IsFileOpen(excelFilePath))
            {
                Console.WriteLine("Excel file is still open. Skipping email...");
                return;
            }

            string currentTop4Hash = GetTop4RowsHash();

            if (currentTop4Hash == lastTop4RowsHash)
            {
                Console.WriteLine("No change detected in top 4 rows. Email not sent.");
                return;
            }

            // File changed in top 4 rows
            lastTop4RowsHash = currentTop4Hash;

            // Optionally, check stability by last modified time
            DateTime lastModified = File.GetLastWriteTime(excelFilePath);
            if ((DateTime.Now - lastModified) < TimeSpan.FromSeconds(5)) // adjust delay if needed
            {
                Console.WriteLine("Waiting for file to stabilize...");
                return;
            }

            // Get recipients
            var recipients = GetEmailListFromExcel();
            if (recipients.Count == 0)
            {
                Console.WriteLine("No recipients found.");
                return;
            }

            // Capture screenshot and send email
            //string imagePath = CaptureGridScreenshot(); // Commented this and next line for sending email and captureing grid view.
            //SendOutlookEmailWithImage(recipients, imagePath);
            //SendSMSNotification("Dyno schedule updated and saved successfully.");
            //Console.WriteLine($"Email sent at {DateTime.Now}");

        }
        private string GetTop4RowsHash()
        {
            

            StringBuilder sb = new StringBuilder();

            try
            {
                using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(stream))
                {
                    foreach (var ws in workbook.Worksheets)
                    {
                        if (ws.Name.Equals("UserData", StringComparison.OrdinalIgnoreCase) || ws.Name.Equals("CompletedTests", StringComparison.OrdinalIgnoreCase) || ws.Name.Equals("ReportUploaded", StringComparison.OrdinalIgnoreCase))
                            continue;

                        int maxRow = Math.Min(4, ws.LastRowUsed()?.RowNumber() ?? 0);

                        for (int row = 1; row <= maxRow; row++)
                        {
                            sb.Append($"Sheet:{ws.Name}|R{row}|");
                            foreach (var cell in ws.Row(row).CellsUsed())
                            {
                                sb.Append(cell.GetValue<string>().Trim());
                                sb.Append("|");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error reading Excel for top 4 rows hash: " + ex.Message);
            }

            return sb.ToString();

        }

        private bool IsFileOpen(string path)
        {
            try
            {
                using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    return false; // Not open
                }
            }
            catch (IOException)
            {
                return true; // Locked
            }
        }

        private List<string> GetEmailListFromExcel()
        {
            List<string> emails = new List<string>();

            try
            {
                using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataSet = reader.AsDataSet();
                    var userSheet = dataSet.Tables.Cast<DataTable>()
                        .FirstOrDefault(t => t.TableName.Equals("UserData", StringComparison.OrdinalIgnoreCase));

                    if (userSheet == null) return emails;

                    for (int i = 1; i < userSheet.Rows.Count; i++)
                    {
                        var email = userSheet.Rows[i][1]?.ToString()?.Trim();
                        if (!string.IsNullOrEmpty(email))
                            emails.Add(email);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading UserData sheet: " + ex.Message);
            }

            return emails;
        }


        private void SendOutlookEmailWithImage(List<string> recipients, string imagePath)
        {
            // Works for outlook email system only this commented part
            //try
            //{
            //    Outlook.Application outlookApp = new Outlook.Application();
            //    Outlook.MailItem mail = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

            //    mail.Subject = "Dyno Schedule Updated - " + Path.GetFileName(excelFilePath);
            //    mail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;

            //    string htmlBody = $@"
            //        <html>
            //        <body>
            //        <p>Hello Team,</p>
            //        <p>The Dyno schedule has been updated.</p>
            //        <p><b>Plymouth Dyno Schedule:</b></p>
            //        <img src='cid:gridshot'>
            //        <p>Regards,<br>Dyno Monitoring System</p>
            //        </body>
            //        </html>";

            //    mail.HTMLBody = htmlBody;

            //    // Add recipients
            //    foreach (var recipient in recipients)
            //        mail.Recipients.Add(recipient);

            //    // Attach image inline
            //    var attach = mail.Attachments.Add(imagePath, Outlook.OlAttachmentType.olByValue, 1, "Grid View");
            //    attach.PropertyAccessor.SetProperty(
            //        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "gridshot");

            //    mail.Send();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Email sending failed: " + ex.Message);
            //}
            //--------------------------------------------------------
            //try
            //{
            //    // ⚠️ Outlook email sending is disabled for TSAC compliance.
            //    // You can integrate Twilio/Fast2SMS here later if needed.

            //    string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "EmailLog.txt");
            //    File.AppendAllText(logPath,
            //        $"[{DateTime.Now}] Email placeholder triggered for recipients: {string.Join(", ", recipients)}\n");

            //    MessageBox.Show("Email notification skipped (Outlook disabled for TSAC compliance).",
            //        "Notification Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Notification logging failed: " + ex.Message);
            //}

        }

        private void SendSMSNotification(string message)
        {
            // Future placeholder for SMS (Twilio, Fast2SMS, etc.)
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SMSLog.txt");
            File.AppendAllText(logPath, $"[{DateTime.Now}] SMS placeholder: {message}\n");
        }

        //private string CaptureGridScreenshot()
        //{
        //    string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GridScreenshot.png");

        //    Bitmap bmp = new Bitmap(dgvSummary.Width, dgvSummary.Height);
        //    dgvSummary.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
        //    bmp.Save(imagePath, System.Drawing.Imaging.ImageFormat.Png);

        //    return imagePath;
        //}

        private void LoadSummary()
        {
            //if (!File.Exists(excelFilePath)) return;

            //DataTable dt = CreateDataTable();

            //try
            //{
            //    using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            //    using (var workbook = new XLWorkbook(stream))
            //    {
            //        foreach (var ws in workbook.Worksheets)
            //        {
            //            if (ws.Name.Equals("UserData",StringComparison.OrdinalIgnoreCase) || ws.Name.Equals("CompletedTests", StringComparison.OrdinalIgnoreCase))
            //                continue;

            //            for (int row = 2; row <= 5; row++)
            //            {
            //                if (ws.Row(row).IsEmpty()) continue;

            //                DataRow dr = dt.NewRow();
            //                FillDataRow(dr, ws, row);
            //                dt.Rows.Add(dr);
            //            }
            //        }
            //    }
            //}
            //catch (IOException ex)
            //{
            //    // Optional: log or show message, but don't crash
            //    Console.WriteLine("Excel file is currently in use: " + ex.Message);
            //    return;
            //}

            //dgvSummary.DataSource = dt;
            //FormatGrid(dgvSummary, summary: true);
            //ApplyDynoColors(dgvSummary);

            if (!File.Exists(excelFilePath)) return;

            DataTable dt = CreateDataTable();

            // HOLD column index in Excel (assumed after Test Owner col 14 => HOLD col 15)
            const int HOLD_COL = 15;

            try
            {
                using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(stream))
                {
                    foreach (var ws in workbook.Worksheets)
                    {
                        if (ws.Name.Equals("UserData", StringComparison.OrdinalIgnoreCase) ||
                            ws.Name.Equals("CompletedTests", StringComparison.OrdinalIgnoreCase) || ws.Name.Equals("ReportsUploaded", StringComparison.OrdinalIgnoreCase))
                            continue;

                        // Top 4 rows = Excel rows 2..5
                        for (int row = 2; row <= 5; row++)
                        {
                            if (ws.Row(row).IsEmpty()) continue;

                            DataRow dr = dt.NewRow();
                            FillDataRow(dr, ws, row);

                            // HOLD detection into helper cols
                            string holdVal = ws.Cell(row, HOLD_COL).GetValue<string>()?.Trim();
                            string reason = ExtractHoldReason(holdVal);

                            if (!string.IsNullOrWhiteSpace(reason))
                            {
                                dr["_HoldFlag"] = "1";
                                dr["_HoldMessage"] = $"DYNO {ws.Name} OFFLINE FOR: {reason}";
                            }
                            else
                            {
                                dr["_HoldFlag"] = "";
                                dr["_HoldMessage"] = "";
                            }

                            dt.Rows.Add(dr);
                        }
                    }
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine("Excel file is currently in use: " + ex.Message);
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine("LoadSummary error: " + ex.Message);
                return;
            }

            // ===== FAST UI UPDATE (NO TOP-TO-BOTTOM DRAW) =====
            BeginGridUpdate(dgvSummary);

            try
            {
                dgvSummary.DataSource = dt;

                if (dgvSummary.Columns.Contains("Hold Dyno For Engineering/ Repair/ Maintenance"))
                {
                    var c = dgvSummary.Columns["Hold Dyno For Engineering/ Repair/ Maintenance"];

                    // Shorter header (viewer-only, Excel untouched)
                    c.HeaderText = "Dyno Hold";

                    // Stop it from expanding
                    c.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    c.Width = 150;                 // adjust if needed
                    c.MinimumWidth = 120;

                    // Wrap header text instead of growing wide
                    c.HeaderCell.Style.WrapMode = DataGridViewTriState.True;
                }
                /* ===== END BLOCK ===== */

                // (Optional) allow ONE sensible column to fill remaining space
                if (dgvSummary.Columns.Contains("Test Script"))
                {
                    dgvSummary.Columns["Test Script"].AutoSizeMode =
                        DataGridViewAutoSizeColumnMode.Fill;
                }

                // Continue with existing logic
                FormatGrid(dgvSummary, summary: true);
                ApplyDynoColors(dgvSummary);
                // Hide helper columns AFTER DataSource bind
                if (dgvSummary.Columns.Contains("_HoldFlag"))
                    dgvSummary.Columns["_HoldFlag"].Visible = false;

                if (dgvSummary.Columns.Contains("_HoldMessage"))
                    dgvSummary.Columns["_HoldMessage"].Visible = false;

                // IMPORTANT: Calling FormatGrid every refresh can be slow if it autosizes.
                // Ideally call FormatGrid ONCE at startup.
                // If your FormatGrid is lightweight, you can keep it. If it causes flicker, comment it out.
                // FormatGrid(dgvSummary, summary: true);

                ApplyDynoColors(dgvSummary);
                // Prevent last column from expanding too much
                if (dgvSummary.Columns.Contains("Hold Dyno For Engineering/ Repair/ Maintenance"))
                {
                    var holdCol = dgvSummary.Columns["Hold Dyno For Engineering/ Repair/ Maintenance"];

                    holdCol.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    holdCol.Width = 160;              // adjust as you like (140–180 works well)
                    holdCol.MinimumWidth = 120;
                    holdCol.Resizable = DataGridViewTriState.True;

                    // Optional: wrap header text so it uses height instead of width
                    holdCol.HeaderCell.Style.WrapMode = DataGridViewTriState.True;
                }

                // Make HOLD rows transparent so banner text is clean
                foreach (DataGridViewRow r in dgvSummary.Rows)
                {
                    if (r.IsNewRow) continue;

                    var flag = r.Cells["_HoldFlag"]?.Value?.ToString();
                    if (flag == "1")
                    {
                        r.DefaultCellStyle.ForeColor = Color.Transparent;
                        r.DefaultCellStyle.SelectionForeColor = Color.Transparent;
                    }
                    else
                    {
                        // restore defaults
                        r.DefaultCellStyle.ForeColor = dgvSummary.DefaultCellStyle.ForeColor;
                        r.DefaultCellStyle.SelectionForeColor = dgvSummary.DefaultCellStyle.SelectionForeColor;
                    }
                }

                // Trigger repaint for the HOLD banners (your Paint handler)
                dgvSummary.Invalidate();
            }
            finally
            {
                EndGridUpdate(dgvSummary);
            }
        }
        private DateTime lastSeenWriteTime = DateTime.MinValue;
        private void LoadDetails(string dynoName, bool preserveSearch = false)
        {
            //if (dynoName.Equals("UserData", StringComparison.OrdinalIgnoreCase)) return;
            if (!File.Exists(excelFilePath)) return;

            detailsTable = CreateDataTable();

            try
            {
                // Open Excel file in read-only mode, allowing shared access
                using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(stream))
                {
                    var ws = workbook.Worksheet(dynoName);

                    int lastRow = ws.LastRowUsed().RowNumber();

                    for (int row = 2; row <= lastRow; row++)
                    {
                        if (ws.Row(row).IsEmpty()) continue;

                        DataRow dr = detailsTable.NewRow();
                        FillDataRow(dr, ws, row);
                        detailsTable.Rows.Add(dr);
                    }
                }
            }
            catch (IOException ex)
            {
                // Optional: log or show message but don't crash
                Console.WriteLine("Excel file is currently in use: " + ex.Message);
                return;
            }

            // Bind to DataGridView
            dgvDetails.DataSource = detailsTable;
            


            FormatGrid(dgvDetails, summary: false);

            // Apply previous search if requested
            if (preserveSearch && !string.IsNullOrEmpty(txtSearch.Text))
                ApplySearch(txtSearch.Text);
        }
        private void btnLoadDyno_Click_1(object sender, EventArgs e)
        {
            if (cmbDyno.SelectedItem == null) return;

            LoadDetails(cmbDyno.SelectedItem.ToString());

            // Switch to details tab
            tabControl1.SelectedTab = tabPage2;
        }

        private void txtSearch_TextChanged_1(object sender, EventArgs e)
        {
            ApplySearch(txtSearch.Text);
            
        }
        

        private void ApplySearch(string filterText)
        {
            

            if (detailsTable == null || detailsTable.Rows.Count == 0) return;

            string escaped = filterText.Trim().Replace("'", "''").ToLower();

            if (string.IsNullOrEmpty(escaped))
            {
                dgvDetails.DataSource = detailsTable;
            }
            else
            {
                DataView dv = new DataView(detailsTable);

                // Make filter case-insensitive by converting columns to lower
                dv.RowFilter = $"CONVERT([Test Number], 'System.String') LIKE '%{escaped}%' OR " +
                               $"CONVERT([Project ID], 'System.String') LIKE '%{escaped}%' OR " +
                               $"CONVERT([Brake Name], 'System.String') LIKE '%{escaped}%'";

                dgvDetails.DataSource = dv;
            }
        }
        private void LoadDynoList()
        {
            if (!File.Exists(excelFilePath)) return;

            cmbDyno.Items.Clear();

            try
            {
                // Open Excel file in read-only mode, allowing shared access
                using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(stream))
                {
                    foreach (var ws in workbook.Worksheets)
                    {
                        if (ws.Name.Equals("UserData", StringComparison.OrdinalIgnoreCase) || ws.Name.Equals("CompletedTests", StringComparison.OrdinalIgnoreCase) || ws.Name.Equals("ReportsUploaded", StringComparison.OrdinalIgnoreCase))
                            continue; // Skip UserData sheet

                        cmbDyno.Items.Add(ws.Name);
                    }
                }

                if (cmbDyno.Items.Count > 0)
                    cmbDyno.SelectedIndex = 0;
            }
            catch (IOException ex)
            {
                // Optional: log or show message but don't crash
                Console.WriteLine("Excel file is currently in use: " + ex.Message);
            }
        }




        private void LoadDetails(string dynoName)
        {
            if (dynoName.Equals("UserData", StringComparison.OrdinalIgnoreCase)) return;
            if (!File.Exists(excelFilePath)) return;

            DataTable dt = CreateDataTable();

            try
            {
                // Open Excel file in read-only mode, allowing shared access
                using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var workbook = new XLWorkbook(stream))
                {
                    var ws = workbook.Worksheet(dynoName);
                    int lastRow = ws.LastRowUsed().RowNumber();

                    for (int row = 2; row <= lastRow; row++)
                    {
                        if (ws.Row(row).IsEmpty()) continue;

                        DataRow dr = dt.NewRow();
                        FillDataRow(dr, ws, row);
                        dt.Rows.Add(dr);
                    }
                }
            }
            catch (IOException ex)
            {
                // Optional: log or show message but don't crash
                Console.WriteLine("Excel file is currently in use: " + ex.Message);
                return;
            }

            dgvDetails.DataSource = dt;

            FormatGrid(dgvDetails, summary: false);
        }


        private DataTable CreateDataTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Dyno");
            dt.Columns.Add("Priority");
            dt.Columns.Add("Duration (hr)");
            dt.Columns.Add("Test Number");
            dt.Columns.Add("Development Code");
            dt.Columns.Add("Sales Code");
            dt.Columns.Add("Reason for test");
            dt.Columns.Add("Test Specification");
            dt.Columns.Add("Project ID");
            dt.Columns.Add("Test Script");
            dt.Columns.Add("Brake Name");
            dt.Columns.Add("Hardware");
            dt.Columns.Add("Requestor");
            dt.Columns.Add("Test Owner");
            dt.Columns.Add("Hold Dyno For Engineering/ Repair/ Maintenance");
            dt.Columns.Add("_HoldFlag");     // "1" or ""
            dt.Columns.Add("_HoldMessage");  // message to display
            return dt;
        }

        private void FillDataRow(DataRow dr, IXLWorksheet ws, int row)
        {
            dr["Dyno"] = ws.Name;
            dr["Priority"] = ws.Cell(row, 2).GetValue<string>();
            dr["Duration (hr)"] = ws.Cell(row, 3).GetValue<string>();
            dr["Test Number"] = ws.Cell(row, 4).GetValue<string>();
            dr["Development Code"] = ws.Cell(row, 5).GetValue<string>();
            dr["Sales Code"] = ws.Cell(row, 6).GetValue<string>();
            dr["Reason for test"] = ws.Cell(row, 7).GetValue<string>();
            dr["Test Specification"] = ws.Cell(row, 8).GetValue<string>();
            dr["Project ID"] = ws.Cell(row, 9).GetValue<string>();
            dr["Test Script"] = ws.Cell(row, 10).GetValue<string>();
            dr["Brake Name"] = ws.Cell(row, 11).GetValue<string>();
            dr["Hardware"] = ws.Cell(row, 12).GetValue<string>();
            dr["Requestor"] = ws.Cell(row, 13).GetValue<string>();
            dr["Test Owner"] = ws.Cell(row, 14).GetValue<string>();
            dr["Hold Dyno For Engineering/ Repair/ Maintenance"] = ws.Cell(row, 15).GetValue<string>();
            const int HOLD_COL = 15;

            string holdVal = ws.Cell(row, HOLD_COL).GetValue<string>()?.Trim();
            if (!string.IsNullOrWhiteSpace(holdVal) && holdVal.StartsWith("Yes", StringComparison.OrdinalIgnoreCase))
            {
                dr["_HoldFlag"] = "1";
                dr["_HoldMessage"] = $"{ws.Name} Dyno DOWN - {holdVal}";
            }
            else
            {
                dr["_HoldFlag"] = "";
                dr["_HoldMessage"] = "";
            }
        }
        
        private void FormatGrid(DataGridView dgv, bool summary)
        {
            //dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            float baseFontSize = 9f;
            if (dgv.Rows.Count > 20) baseFontSize = 9f;
            if (dgv.Rows.Count > 40) baseFontSize = 9f;

            Font gridFont = new Font("Segoe UI", baseFontSize, FontStyle.Regular);
            dgv.Font = gridFont;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            
            dgv.RowTemplate.Height = 8;
            dgv.AllowUserToResizeRows = false;

            foreach (DataGridViewColumn col in dgv.Columns)
            {
                col.HeaderCell.Style.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                col.MinimumWidth = 50;
                col.Resizable = DataGridViewTriState.True;
            }
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.Height = (int)(gridFont.GetHeight() * 0.5); // multiplier for padding
            }
            // First 4 columns = small fixed width
            dgv.Columns[0].Width = 80;  // Dyno
            dgv.Columns[1].Width = 60;  // Priority
            dgv.Columns[2].Width = 100;  // Duration
            dgv.Columns[3].Width = 100; // Test Number
            dgv.Columns[4].Width = 130; // Delvelopment code
            dgv.Columns[5].Width = 100; // Sales code

            // Remaining columns auto fill
            //dgv.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; // Delvelopment code
            //dgv.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; // Sales code
            dgv.Columns[6].Width = 130; // Reason for Test
            dgv.Columns[7].Width = 130; // Test Specification
            dgv.Columns[9].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; // Project ID
            dgv.Columns[10].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; // Test Script
            //dgv.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; // Brake Name
            dgv.Columns[8].Width = 100;
            dgv.Columns[13].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; // Requestor

            dgv.RowPostPaint -= Dgv_RowPostPaint; // remove duplicate handlers
            dgv.RowPostPaint += Dgv_RowPostPaint;
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //float baseFontSize = 9f;
            //if (dgv.Rows.Count > 20) baseFontSize = 8f;
            //if (dgv.Rows.Count > 40) baseFontSize = 7f;

            //Font gridFont = new Font("Segoe UI", baseFontSize, FontStyle.Regular);
            //dgv.Font = gridFont;
            //dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //dgv.RowTemplate.Height = 30;
            //dgv.AllowUserToResizeRows = false;

            //foreach (DataGridViewColumn col in dgv.Columns)
            //{
            //    col.HeaderCell.Style.Font = new Font("Segoe UI", gridFont.Size, FontStyle.Bold);
            //    col.SortMode = DataGridViewColumnSortMode.NotSortable;
            //    col.MinimumWidth = 50;
            //    col.Resizable = DataGridViewTriState.True;
            //}

            //dgv.Columns["Dyno"].Width = 80;
            //dgv.Columns["Priority"].Width = 60;
            //dgv.Columns["Duration (hr)"].Width = 80;
            //dgv.Columns["Test Number"].Width = 100;
            //dgv.Columns["Development Code"].Width = 100;
            //dgv.Columns["Sales Code"].Width = 100;

        }
        private void Dgv_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var dgv = sender as DataGridView;
            if (dgv == null) return;

            if (e.RowIndex < 0 || e.RowIndex >= dgv.Rows.Count) return;

            string currentDyno = dgv.Rows[e.RowIndex].Cells["Dyno"].Value?.ToString() ?? "";

            string nextDyno = "";
            if (e.RowIndex < dgv.Rows.Count - 1)
                nextDyno = dgv.Rows[e.RowIndex + 1].Cells["Dyno"].Value?.ToString() ?? "";

            if (currentDyno != nextDyno)
            {
                // Use GetRowDisplayRectangle to get row bounds
                Rectangle rowRect = dgv.GetRowDisplayRectangle(e.RowIndex, true);
                using (Pen pen = new Pen(Color.Black, 2)) // bold line
                {
                    int y = rowRect.Bottom - 1;
                    e.Graphics.DrawLine(pen, rowRect.Left, y, rowRect.Right, y);
                }
            }
        }
        private void ApplyDynoColors(DataGridView dgv)
        {
            //Color[] dynoColors = { Color.LightBlue, Color.LightYellow, Color.LightGreen, Color.LightPink, Color.LightGray };
            //int colorIndex = -1;
            //string lastDyno = "";

            //foreach (DataGridViewRow row in dgv.Rows)
            //{
            //    if (row.Cells["Dyno"].Value == null) continue;

            //    string currentDyno = row.Cells["Dyno"].Value.ToString();

            //    if (currentDyno != lastDyno)
            //    {
            //        colorIndex++;
            //        lastDyno = currentDyno;
            //    }

            //    row.DefaultCellStyle.BackColor = dynoColors[colorIndex % dynoColors.Length];

            //    // Hardware True/False
            //    if (row.Cells["Hardware"].Value?.ToString().Equals("True", StringComparison.OrdinalIgnoreCase) == true)
            //    {
            //        row.Cells["Hardware"].Style.BackColor = Color.LightGreen;
            //    }
            //    else if (row.Cells["Hardware"].Value?.ToString().Equals("False", StringComparison.OrdinalIgnoreCase) == true)
            //    {
            //        row.Cells["Hardware"].Style.BackColor = Color.LightCoral;
            //    }
            //if (i == dgv.Rows.Count - 1 || (dgv.Rows[i + 1].Cells["Dyno"].Value?.ToString() != currentDyno))
            //{
            //    row.DefaultCellStyle.Padding = new Padding(0, 0, 0, 4); // space for border
            //    row.DefaultCellStyle.SelectionBackColor = row.DefaultCellStyle.BackColor;
            //    row.DividerHeight = 4; // thickness of the separation line
            //    row.DividerColor = Color.Black; // bold line color
            //}

            Color[] dynoColors = { Color.LightBlue, Color.LightYellow, Color.LightGreen, Color.LightPink, Color.LightGray };
            int colorIndex = -1;
            string lastDyno = "";

            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                var row = dgv.Rows[i];
                if (row.Cells["Dyno"].Value == null) continue;

                string currentDyno = row.Cells["Dyno"].Value.ToString();
                if (currentDyno != lastDyno)
                {
                    colorIndex++;
                    lastDyno = currentDyno;
                }

                row.DefaultCellStyle.BackColor = dynoColors[colorIndex % dynoColors.Length];
                row.DefaultCellStyle.SelectionBackColor = row.DefaultCellStyle.BackColor;

                if (row.Cells["Hardware"].Value?.ToString().Equals("True", StringComparison.OrdinalIgnoreCase) == true)
                    row.Cells["Hardware"].Style.BackColor = Color.LightGreen;
                else if (row.Cells["Hardware"].Value?.ToString().Equals("False", StringComparison.OrdinalIgnoreCase) == true)
                    row.Cells["Hardware"].Style.BackColor = Color.LightCoral;

                if (i == dgv.Rows.Count - 1 || (dgv.Rows[i + 1].Cells["Dyno"].Value?.ToString() != currentDyno))
                {
                    row.DividerHeight = 3;
                    //row.DividerColor = Color.Black;
                    row.DefaultCellStyle.Padding = new Padding(0, 0, 0, 4);
                }
                else
                {
                    row.DividerHeight = 1;
                    //row.DividerColor = Color.LightGray;
                    row.DefaultCellStyle.Padding = new Padding(0);
                }
            }
        
        }

        private void dgvSummary_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //if (e.ColumnIndex < 0 || e.ColumnIndex >= dgvDetails.Columns.Count)
            //    return;

            //// Check if we are formatting the "Hardware" column
            //var column = dgvSummary.Columns[e.ColumnIndex];
            //if (column != null && column.Name.Equals("Hardware", StringComparison.OrdinalIgnoreCase) && e.Value != null)
            //{
            //    string cellValue = e.Value.ToString().Trim().ToLower();
            //    e.Value = (cellValue == "TRUE") ? "True" : "False";
            //    e.FormattingApplied = true;
            //}
        }
        private bool TryGetHoldReason(string holdCell, out string reason)
        {
            reason = "";
            if (string.IsNullOrWhiteSpace(holdCell)) return false;

            string s = holdCell.Trim();

            // accept: "Yes - Repair", "Yes-Repair", "YES - Engineering"
            if (!s.StartsWith("Yes", StringComparison.OrdinalIgnoreCase)) return false;

            int dash = s.IndexOf('-');
            if (dash >= 0 && dash < s.Length - 1)
                reason = s.Substring(dash + 1).Trim();
            else
                reason = "Hold";

            return true;
        }

        private void ShowHoldBanner(string message)
        {
            lblHoldBanner.Text = message;

            pnlHoldBanner.Visible = true;
            pnlHoldBanner.BringToFront();

            // Shrink grid so banner doesn't cover the top rows
            dgvSummary.Top = dgvSummaryOriginalTop + pnlHoldBanner.Height + 6;
            dgvSummary.Height = dgvSummaryOriginalHeight - (pnlHoldBanner.Height + 6);
        }

        private void HideHoldBanner()
        {
            pnlHoldBanner.Visible = false;

            // Restore grid size
            dgvSummary.Top = dgvSummaryOriginalTop;
            dgvSummary.Height = dgvSummaryOriginalHeight;
        }
        private void DgvSummary_PaintHoldBanners(object sender, PaintEventArgs e)
        {
            var dgv = dgvSummary;
            if (dgv.Rows.Count == 0) return;

            // Walk rows and find dyno blocks
            int i = 0;
            while (i < dgv.Rows.Count)
            {
                if (dgv.Rows[i].IsNewRow) break;

                string dyno = dgv.Rows[i].Cells["Dyno"].Value?.ToString() ?? "";
                if (string.IsNullOrWhiteSpace(dyno)) { i++; continue; }

                // Find the end of this dyno block (usually 4 rows, but safe)
                int start = i;
                int end = i;
                while (end + 1 < dgv.Rows.Count &&
                       (dgv.Rows[end + 1].Cells["Dyno"].Value?.ToString() ?? "") == dyno)
                {
                    end++;
                }

                // Find first HOLD row inside this dyno block
                int holdStart = -1;
                string holdMsg = null;

                for (int r = start; r <= end; r++)
                {
                    var flag = dgv.Rows[r].Cells["_HoldFlag"].Value?.ToString();
                    if (flag == "1")
                    {
                        holdStart = r;
                        holdMsg = dgv.Rows[r].Cells["_HoldMessage"].Value?.ToString();
                        break;
                    }
                }

                // If HOLD found, draw banner from holdStart to end (to occupy that remaining space)
                if (holdStart != -1 && !string.IsNullOrWhiteSpace(holdMsg))
                {
                    Rectangle rectStart = dgv.GetRowDisplayRectangle(holdStart, true);
                    Rectangle rectEnd = dgv.GetRowDisplayRectangle(end, true);

                    // If rows are not visible (scrolled off), skip
                    if (rectStart.Height > 0 && rectEnd.Height > 0)
                    {
                        int left = dgv.RowHeadersVisible ? dgv.RowHeadersWidth : 0;
                        int top = rectStart.Top;
                        int width = dgv.DisplayRectangle.Width - 1;
                        int height = (rectEnd.Bottom - rectStart.Top);

                        Rectangle bannerRect = new Rectangle(left + 1, top + 1, width - 2, height - 2);

                        using (Brush b = new SolidBrush(Color.Gold))
                            e.Graphics.FillRectangle(b, bannerRect);

                        using (Pen p = new Pen(Color.OrangeRed, 2))
                            e.Graphics.DrawRectangle(p, bannerRect);

                        using (Font f = new Font("Segoe UI", 14, FontStyle.Bold))
                        {
                            TextRenderer.DrawText(
                                e.Graphics,
                                "⚠ " + holdMsg,
                                f,
                                bannerRect,
                                Color.Black,
                                TextFormatFlags.HorizontalCenter |
                                TextFormatFlags.VerticalCenter |
                                TextFormatFlags.WordBreak
                            );
                        }
                    }
                }

                i = end + 1;
            }
        }



    }


}

