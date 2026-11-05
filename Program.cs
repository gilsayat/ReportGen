using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ClosedXML.Excel;
using HtmlAgilityPack;

namespace ReportAutomator
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public class ColorSchemePreset
    {
        public string Name { get; set; }
        public Color MainBarColor { get; set; }
        public Color SubBarColor { get; set; }
        public override string ToString() => Name;
    }

    public class AccountingRow
    {
        public string EmployeeId { get; set; }
        public string PhotographerName { get; set; }
        public string CaptureType { get; set; }
        public string Plu { get; set; }
        public int Images { get; set; }
        public int SoldCore { get; set; }
        public int SoldUpsell { get; set; }
        public double TotalAmount { get; set; }
        public string Category { get; set; }
    }

    public class OrderRow
    {
        public string OrderId { get; set; }
        public string CabinNo { get; set; }
        public string GuestName { get; set; }
        public string FileName { get; set; }
        public string EmployeeId { get; set; }
        public string PhotographerName { get; set; }
        public string CaptureType { get; set; }
        public string MediaType { get; set; }
        public string OrderStatus { get; set; }
        public string ChannelId { get; set; }
        public string VoyageNumber { get; set; }
    }

    public class ManifestRow
    {
        public string CabinNum { get; set; }
        public int Quantity { get; set; }
        public string ItemDescription { get; set; }
        public string GuestName { get; set; }
    }

    public class IrisSummary
    {
        public int PremiumPkgQty { get; set; }
        public double PremiumPkgAmt { get; set; }
        public int CanvasQty { get; set; }
        public double CanvasAmt { get; set; }
        public int DigitalUsbQty { get; set; }
        public double DigitalUsbAmt { get; set; }
        public int Print10x10Qty { get; set; }
        public double Print10x10Amt { get; set; }
        public int Print8x10Qty { get; set; }
        public double Print8x10Amt { get; set; }
        public int Print10x13Qty { get; set; }
        public double Print10x13Amt { get; set; }
        public int AddOnQty { get; set; }
        public double AddOnAmt { get; set; }

        public int TotalQty => PremiumPkgQty + CanvasQty + DigitalUsbQty + Print10x10Qty + Print8x10Qty + Print10x13Qty + AddOnQty;
        public double TotalAmt => Math.Round(PremiumPkgAmt + CanvasAmt + DigitalUsbAmt + Print10x10Amt + Print8x10Amt + Print10x13Amt + AddOnAmt, 2);
    }

    public class PivotRenderLine
    {
        public int LineType { get; set; } // 1 = HeaderBar, 2 = PhotographerGroup, 3 = CaptureRow, 4 = GrandTotal
        public string Label { get; set; }
        public string Col1 { get; set; }
        public string Col2 { get; set; }
        public string Col3 { get; set; }
        public string Col4 { get; set; }
    }

    public class MainForm : Form
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool DestroyIcon(IntPtr handle);

        private bool _isDarkMode = false;

        private string _accountingFilePath = "";
        private string _orderFilePath = "";
        private string _manifestFilePath = "";
        private string _fssFilePath = "";
        private string _lastOutputFolder = "";

        private List<AccountingRow> _accountingRows = new List<AccountingRow>();
        private List<OrderRow> _rawOrderRows = new List<OrderRow>();
        private List<OrderRow> _orderRowsForDataTab = new List<OrderRow>();
        private List<OrderRow> _orderRowsForPhotoSales = new List<OrderRow>();
        private HashSet<string> _excludedOrderStatuses = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Refund", "InCart" };
        private HashSet<string> _excludedMediaTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "OB02", "OB03", "OB04", "OB05", "OB2", "OB3", "OB4", "OB5" };
        private Button _btnFilterOrders;
        private Font _gridHeaderFont = new Font("Segoe UI Semibold", 12f, FontStyle.Bold, GraphicsUnit.Pixel);
        private List<ManifestRow> _manifestRows = new List<ManifestRow>();
        private IrisSummary _irisSummary = new IrisSummary();

        // UI Containers
        private Panel _headerCard;
        private Panel _grLogoBadge;
        private Label _lblSubtitle;
        private Label _lblVoyageTag;
        private TextBox _txtVoyage;
        private Label _lblThemeTag;
        private ComboBox _cmbTheme;
        private Button _btnCustomColor;
        private Panel _pnlColorPreview;
        private Button _btnDarkToggle;

        private Panel _dropCard;
        private Label _lblDropTitle;
        private Label _lblDropHint;
        private Button _btnBrowseFiles;
        private Button _btnClearFiles;

        private TableLayoutPanel _statusTable;
        private Label _badgeAcc;
        private Label _badgeOrd;
        private Label _badgeMan;
        private Label _badgeFss;

        private Panel _previewCard;
        private TabControl _tabPreview;
        private DataGridView _dgvCapture;
        private DataGridView _dgvSeuss;
        private DataGridView _dgvIris;
        private DataGridView _dgvPivot;
        private DataGridView _dgvOrderPhoto;
        private DataGridView _dgvOrderSales;

        private Panel _footerCard;
        private CheckBox _chkExportCsvs;
        private Button _btnCopyTab;
        private Button _btnOpenFolder;
        private Button _btnGenerateAll;
        private Label _lblFooterStatus;
        private Label _lblCreatedBy;

        private List<ColorSchemePreset> _presets = new List<ColorSchemePreset>
        {
            new ColorSchemePreset { Name = "Orchid Purple (Default)", MainBarColor = ColorTranslator.FromHtml("#D86DCD"), SubBarColor = ColorTranslator.FromHtml("#F2CEEF") },
            new ColorSchemePreset { Name = "iPadOS Blue", MainBarColor = ColorTranslator.FromHtml("#007AFF"), SubBarColor = ColorTranslator.FromHtml("#D6EAFF") },
            new ColorSchemePreset { Name = "Emerald Green", MainBarColor = ColorTranslator.FromHtml("#34C759"), SubBarColor = ColorTranslator.FromHtml("#DCF7E3") },
            new ColorSchemePreset { Name = "Sunset Amber", MainBarColor = ColorTranslator.FromHtml("#FF9500"), SubBarColor = ColorTranslator.FromHtml("#FFEAD0") },
            new ColorSchemePreset { Name = "Crimson Rose", MainBarColor = ColorTranslator.FromHtml("#FF2D55"), SubBarColor = ColorTranslator.FromHtml("#FFD9E0") },
            new ColorSchemePreset { Name = "Midnight Slate", MainBarColor = ColorTranslator.FromHtml("#48484A"), SubBarColor = ColorTranslator.FromHtml("#E5E5EA") },
            new ColorSchemePreset { Name = "Royal Indigo", MainBarColor = ColorTranslator.FromHtml("#5856D6"), SubBarColor = ColorTranslator.FromHtml("#E2E1FF") }
        };

        private Color _activeMainColor = ColorTranslator.FromHtml("#D86DCD");
        private Color _activeSubColor = ColorTranslator.FromHtml("#F2CEEF");

        private static readonly Dictionary<string, string> PluCategoryMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "1ST CASUAL HORIZON RESTAURANT", "CC" },
            { "1ST CASUAL VISTA RESTAURANT", "CC" },
            { "1ST ELEGANT HORIZON RESTAURANT", "CC" },
            { "1ST ELEGANT VISTA RESTAURANT", "CC" },
            { "2ND CABO SAN LUCAS GANGWAY", "CC" },
            { "2ND CABO SL CLASSICS", "CC" },
            { "2ND CABO SL WHEEL", "CC" },
            { "2ND CASUAL HORIZON RESTAURANT", "CC" },
            { "2ND CASUAL VISTA RESTAURANT", "CC" },
            { "2ND ELEGANT HORIZON RESTAURANT", "CC" },
            { "2ND ELEGANT VISTA RESTAURANT", "CC" },
            { "3RD CASUAL HORIZON RESTAURANT", "CC" },
            { "3RD CASUAL VISTA RESTAURANT", "CC" },
            { "4TH CASUAL BOTH DINING", "CC" },
            { "4TH CASUAL HORIZON RESTAURANT", "CC" },
            { "4TH CASUAL VISTA RESTAURANT", "CC" },
            { "80'S DISCO LIGHT", "PP" },
            { "80'S GRAFFITI", "PP" },
            { "80'S NIGHT", "PP" },
            { "80'S NIGHT HORIZON", "CC" },
            { "80'S STEPBACK IN TIME", "PP" },
            { "80S BACK TO FUTURE", "PP" },
            { "ALCHEMY", "PP" },
            { "ALL ABOUT WHITE 1STELEGANT", "PP" },
            { "ALL ABOUT WHITE 2NDELEGANT", "PP" },
            { "ALL ABOUT WHITE CASUAL", "PP" },
            { "BACK TO THE FUTURE", "PP" },
            { "BALLANDRA BEACH", "PP" },
            { "BEACH GATE", "PP" },
            { "BEACH SCENE", "PP" },
            { "BEACH SUNSET", "PP" },
            { "BETH DUTTON", "PP" },
            { "BIG MOON BEACH", "PP" },
            { "BLUE HAWAII", "PP" },
            { "BLUE IGUANA", "PP" },
            { "BLUE WAVE", "PP" },
            { "BRUNCH", "CC" },
            { "CABO SAN LUCAS", "PP" },
            { "CABO SAN LUCAS CLASSICS", "CC" },
            { "CABO SAN LUCAS GANGWAY", "CC" },
            { "CABO SAN LUCAS VILLAGE", "PP" },
            { "CABO SAN LUCAS WHEEL", "CC" },
            { "CANDID DECK", "CC" },
            { "CARIBBEAN BAR", "PP" },
            { "CARIBBEAN NIGHT", "PP" },
            { "CARRIBEAN LIFE", "PP" },
            { "CARVED STAIRCASE", "PP" },
            { "CELESTIAL GARDEN", "PP" },
            { "CHIC HALLWAY", "PP" },
            { "CHOCOLATE BOKEH", "PP" },
            { "DAY WAKE", "PP" },
            { "DECK PARTY", "CC" },
            { "DECK RAIL", "CC" },
            { "DIAMOND RATURE", "PP" },
            { "DOCK ON THE BAY", "PP" },
            { "DR SEUSS BREAKFAST", "CC" },
            { "DR SEUSS CAT IN THE HAT", "PP" },
            { "DR SEUSS SAM I AM", "PP" },
            { "DR SEUSS THING ONE & THING TWO", "PP" },
            { "DUCKY LIFEGUARD", "CC" },
            { "DUCKY OFFICER", "CC" },
            { "DUCKY SHERIFF", "CC" },
            { "DUTTON RANCH", "PP" },
            { "ELEGANT STAIRS", "PP" },
            { "EMBARKATION", "CC" },
            { "ENSENADA", "PP" },
            { "ENSENADA CLASSICS", "CC" },
            { "ENSENADA GANGWAY", "CC" },
            { "ENSENADA TERRACE", "PP" },
            { "ENSENADA WHEEL", "CC" },
            { "FIREWORKS", "PP" },
            { "FLEUR DE LIS", "PP" },
            { "FOREST SHADOW", "PP" },
            { "FRAME ROOM", "PP" },
            { "FUCHSIA FUSION", "PP" },
            { "GANGWAY PIRATE", "CC" },
            { "GOLD GLIMMER", "PP" },
            { "GOLDEN BOKEH ELEGANT", "PP" },
            { "GRAND ROYALE_", "PP" },
            { "GRAND STAIRCASE TITANIC", "PP" },
            { "GRAND TURK BEACH", "PP" },
            { "GREEN SCREEN", "PP" },
            { "HALLOWEEN FOREST", "PP" },
            { "HIGH SEA SAILING", "PP" },
            { "HOT AIR BALLOONS", "PP" },
            { "INDIAN", "CC" },
            { "ISLAND ARUBA", "PP" },
            { "ISLAND SUNSET", "PP" },
            { "JENNIFER", "PP" },
            { "KICKIN BEACH", "PP" },
            { "L.A SKYLINE", "PP" },
            { "LA PAZ", "PP" },
            { "LA PAZ CLASSICS", "CC" },
            { "LA PAZ GANGWAY", "CC" },
            { "LA PAZ WHEEL", "CC" },
            { "LAZY DAY", "PP" },
            { "LIGHTHOUSE SUNSET", "PP" },
            { "LONG BEACH", "PP" },
            { "LOS MUERTOS", "CC" },
            { "MARBLE MYSTIQUE ELEGANT", "PP" },
            { "MARIACHI", "CC" },
            { "MAZATLAN", "PP" },
            { "MAZATLAN CLASSICS", "CC" },
            { "MAZATLAN GANGWAY", "CC" },
            { "MAZATLAN WHEEL", "PP" },
            { "MEXICAN DANCER", "CC" },
            { "MOON & STARS ELEGANT", "PP" },
            { "MOUNTAIN WATERFALL", "PP" },
            { "NEW MOON", "PP" },
            { "NIGHT PARADISE", "PP" },
            { "NIGHT WAKE", "PP" },
            { "NIGHT WATERFALLS", "PP" },
            { "OCEAN DECK", "PP" },
            { "OPTIC RED FLOWER", "PP" },
            { "PALACE REFLECTING POND", "PP" },
            { "PARADISE ISLAND", "PP" },
            { "PARROTS", "CC" },
            { "PIANO CASUAL", "PP" },
            { "PIANO ELEGANT 1ST", "PP" },
            { "PIANO ELEGANT 2ND", "PP" },
            { "PIANO KEYS", "PP" },
            { "PIRATE ROAMING", "CC" },
            { "PIRATE VISTA", "CC" },
            { "PIRATESHIP", "PP" },
            { "PROGRESSO", "PP" },
            { "PUERTO VALLARTA", "PP" },
            { "PUERTO VALLARTA CLASSIC", "CC" },
            { "PUERTO VALLARTA GANGWAY", "CC" },
            { "RAIL EMBARKATION", "CC" },
            { "RED BOUDOIR CASUAL", "PP" },
            { "RED BOUDOIR ELEGANT", "PP" },
            { "REFLECTED MOON", "PP" },
            { "ROSE TRUCK", "PP" },
            { "SAIL AWAY", "CC" },
            { "SAILBOAT", "PP" },
            { "SAILOR GIRL", "CC" },
            { "SEARCH LIGHT BEAM", "PP" },
            { "SEASIDE MYSTIQUE", "PP" },
            { "SEQUIN AQUA", "PP" },
            { "SEQUIN BLUE ROYAL", "PP" },
            { "SEQUIN GOLD", "PP" },
            { "SEQUIN MATTE GOLD", "PP" },
            { "SEQUIN RED CHERRY", "PP" },
            { "SEQUIN ROSEGOLD", "PP" },
            { "SEQUIN SILVER", "PP" },
            { "SEQUIN TEAL", "PP" },
            { "SEQUIN WHITE", "PP" },
            { "SHIP LIGHT BEAM", "PP" },
            { "SHIPS WHEEL", "CC" },
            { "SHIPS WHEEL ON DECK", "CC" },
            { "SHIPSBOW ELEGANT", "PP" },
            { "SKY ROPES", "CC" },
            { "SKYLINE MOUNTAIN", "PP" },
            { "SKYRIDE", "CC" },
            { "SPECIAL EVENT", "PP" },
            { "SPECIAL EVENT 2", "PP" },
            { "SPECIAL EVENTS", "CC" },
            { "SQUARE COLLAGE", "PP" },
            { "STAIRS CASUAL", "PP" },
            { "STAIRS ELEGANT 1ST", "PP" },
            { "STAIRS ELEGANT 2ND", "PP" },
            { "STANDING CASUAL", "PP" },
            { "STANDING ELEGANT 1ST", "PP" },
            { "STANDING ELEGANT 2ND", "PP" },
            { "SUN & SAND CASUAL", "PP" },
            { "SUN WASHED GRAND STAIRS", "PP" },
            { "SUNSET HUTS", "PP" },
            { "SUNSET OVER ATLANTIC", "PP" },
            { "THE TOWEL ELEPHANT", "CC" },
            { "TROPICAL BEACH", "PP" },
            { "TROPICAL COAST", "PP" },
            { "TROPICAL ISLAND SHIP", "PP" },
            { "UMBRELLA STREET", "PP" },
            { "UNCLE SAM", "CC" },
            { "UNICORN SEQUIN", "PP" },
            { "WATER SLIDES", "CC" },
            { "WATERFALLS", "PP" },
            { "WHITE MARBLE HALLWAY", "PP" },
            { "WHITE STAIRCASE ELEGANT", "PP" }
        };

        public MainForm()
        {
            // CRITICAL: AutoScaleMode.None prevents WinForms from inflating fonts while keeping pixel coordinates fixed!
            AutoScaleMode = AutoScaleMode.None;
            Font = new Font("Segoe UI", 9f, FontStyle.Regular);
            Text = "Gil ReportGen v2.0";
            ClientSize = new Size(1160, 760);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = true;
            StartPosition = FormStartPosition.CenterScreen;
            DoubleBuffered = true;
            AllowDrop = true;

            Icon = CreateGrWindowIcon();

            DragEnter += MainForm_DragEnter;
            DragDrop += MainForm_DragDrop;

            BuildUI();
            ApplyTheme();
        }

        private Icon CreateGrWindowIcon()
        {
            using var bmp = new Bitmap(64, 64, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                g.Clear(Color.Transparent);

                var rect = new Rectangle(2, 2, 60, 60);
                using var path = GetRoundedRectPath(rect, 14);
                using var brush = new LinearGradientBrush(rect, Color.FromArgb(0, 122, 255), Color.FromArgb(88, 86, 214), LinearGradientMode.Vertical);
                g.FillPath(brush, path);

                using var font = new Font("Segoe UI", 24f, FontStyle.Bold, GraphicsUnit.Pixel);
                var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                g.DrawString("GR", font, Brushes.White, new RectangleF(0, 1, 64, 64), sf);
            }
            IntPtr hIcon = bmp.GetHicon();
            Icon cloned = (Icon)Icon.FromHandle(hIcon).Clone();
            DestroyIcon(hIcon);
            return cloned;
        }

        private void BuildUI()
        {
            // All fonts use GraphicsUnit.Pixel so 125%/150% Windows Display Scaling NEVER overflows any box or button
            var fontSubBold = new Font("Segoe UI Semibold", 13f, FontStyle.Bold, GraphicsUnit.Pixel);
            var fontLabelBold = new Font("Segoe UI Semibold", 12f, FontStyle.Bold, GraphicsUnit.Pixel);
            var fontInput = new Font("Segoe UI Semibold", 13f, FontStyle.Bold, GraphicsUnit.Pixel);
            var fontCombo = new Font("Segoe UI", 12.5f, FontStyle.Regular, GraphicsUnit.Pixel);
            var fontDropTitle = new Font("Segoe UI Semibold", 15f, FontStyle.Bold, GraphicsUnit.Pixel);
            var fontRegular = new Font("Segoe UI", 12f, FontStyle.Regular, GraphicsUnit.Pixel);
            var fontItalic = new Font("Segoe UI Semibold", 11.5f, FontStyle.Italic, GraphicsUnit.Pixel);
            var fontTab = new Font("Segoe UI Semibold", 12f, FontStyle.Bold, GraphicsUnit.Pixel);

            // 1. Top Control Bar (1124px wide, spacious layout)
            _headerCard = CreateCardPanel(new Rectangle(18, 14, 1124, 58));

            _grLogoBadge = new Panel
            {
                Location = new Point(14, 11),
                Size = new Size(36, 36)
            };
            _grLogoBadge.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                var r = new Rectangle(0, 0, _grLogoBadge.Width - 1, _grLogoBadge.Height - 1);
                using var path = GetRoundedRectPath(r, 9);
                using var br = new LinearGradientBrush(r, Color.FromArgb(0, 122, 255), Color.FromArgb(88, 86, 214), LinearGradientMode.Vertical);
                e.Graphics.FillPath(br, path);
                using var f = new Font("Segoe UI", 14f, FontStyle.Bold, GraphicsUnit.Pixel);
                var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                e.Graphics.DrawString("GR", f, Brushes.White, new RectangleF(0, 0, _grLogoBadge.Width, _grLogoBadge.Height), sf);
            };

            _lblSubtitle = new Label
            {
                Text = "Automated Multi-Tab Excel  •  Accounting PDF  •  Sales.jpg",
                Font = fontSubBold,
                Location = new Point(60, 11),
                Size = new Size(410, 36),
                TextAlign = ContentAlignment.MiddleLeft
            };

            _lblVoyageTag = new Label
            {
                Text = "VOYAGE:",
                Font = fontLabelBold,
                Location = new Point(478, 11),
                Size = new Size(64, 36),
                TextAlign = ContentAlignment.MiddleRight
            };

            _txtVoyage = new TextBox
            {
                Text = "PO102426",
                Font = fontInput,
                Location = new Point(548, 16),
                Size = new Size(105, 26),
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = HorizontalAlignment.Center
            };

            _lblThemeTag = new Label
            {
                Text = "COLOR THEME:",
                Font = fontLabelBold,
                Location = new Point(664, 11),
                Size = new Size(98, 36),
                TextAlign = ContentAlignment.MiddleRight
            };

            _cmbTheme = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = fontCombo,
                Location = new Point(768, 16),
                Size = new Size(185, 26)
            };
            foreach (var p in _presets) _cmbTheme.Items.Add(p);
            _cmbTheme.SelectedIndex = 0;
            _cmbTheme.SelectedIndexChanged += (s, e) =>
            {
                if (_cmbTheme.SelectedItem is ColorSchemePreset preset)
                {
                    _activeMainColor = preset.MainBarColor;
                    _activeSubColor = preset.SubBarColor;
                    _pnlColorPreview.Invalidate();
                }
            };

            _pnlColorPreview = new Panel
            {
                Location = new Point(961, 15),
                Size = new Size(28, 28),
                Cursor = Cursors.Hand
            };
            _pnlColorPreview.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                int half = _pnlColorPreview.Width / 2;
                using var b1 = new SolidBrush(_activeMainColor);
                using var b2 = new SolidBrush(_activeSubColor);
                e.Graphics.FillRectangle(b1, 0, 0, half, _pnlColorPreview.Height);
                e.Graphics.FillRectangle(b2, half, 0, _pnlColorPreview.Width - half, _pnlColorPreview.Height);
                using var pen = new Pen(Color.FromArgb(160, 160, 165), 1);
                e.Graphics.DrawRectangle(pen, 0, 0, _pnlColorPreview.Width - 1, _pnlColorPreview.Height - 1);
            };
            _pnlColorPreview.Click += BtnCustomColor_Click;

            _btnCustomColor = CreatePillButton("Custom", new Point(996, 14), new Size(74, 30), false);
            _btnCustomColor.Click += BtnCustomColor_Click;

            _btnDarkToggle = CreatePillButton("☾", new Point(1076, 14), new Size(34, 30), false);
            _btnDarkToggle.Click += (s, e) =>
            {
                _isDarkMode = !_isDarkMode;
                _btnDarkToggle.Text = _isDarkMode ? "☀" : "☾";
                ApplyTheme();
            };

            _headerCard.Controls.AddRange(new Control[] {
                _grLogoBadge, _lblSubtitle, _lblVoyageTag, _txtVoyage, _lblThemeTag, _cmbTheme, _pnlColorPreview, _btnCustomColor, _btnDarkToggle
            });

            // 2. Drag & Drop Zone + 4 Equal Status Pills
            _dropCard = CreateCardPanel(new Rectangle(18, 84, 1124, 132));
            _dropCard.AllowDrop = true;
            _dropCard.DragEnter += MainForm_DragEnter;
            _dropCard.DragDrop += MainForm_DragDrop;

            _lblDropTitle = new Label
            {
                Text = "Drag & Drop Your 4 Cruise Files Here",
                Font = fontDropTitle,
                Location = new Point(18, 12),
                Size = new Size(550, 22),
                UseMnemonic = false
            };

            _lblDropHint = new Label
            {
                Text = "Auto-detects AccountReportStatistics (.xls), OrderReport (.xls), Manifest (.xlsx), and FSS Item Sales (.csv)",
                Font = fontRegular,
                Location = new Point(18, 37),
                Size = new Size(670, 20),
                UseMnemonic = false
            };

            _btnFilterOrders = CreatePillButton("Media / Status Filters", new Point(696, 15), new Size(160, 36), false);
            _btnFilterOrders.Click += (s, e) => ShowOrderFilterChecklistDialog();

            _btnBrowseFiles = CreatePillButton("Browse Files...", new Point(864, 15), new Size(146, 36), true);
            _btnBrowseFiles.Click += BtnBrowseFiles_Click;

            _btnClearFiles = CreatePillButton("Reset", new Point(1018, 15), new Size(90, 36), false);
            _btnClearFiles.Click += (s, e) => ClearAllData();

            _statusTable = new TableLayoutPanel
            {
                Location = new Point(14, 66),
                Size = new Size(1096, 52),
                ColumnCount = 4,
                RowCount = 1,
                BackColor = Color.Transparent
            };
            _statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            _statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            _statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            _statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            _statusTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            _badgeAcc = CreateStatusBadge("1. Accounting: Waiting...");
            _badgeOrd = CreateStatusBadge("2. Order Report: Waiting...");
            _badgeMan = CreateStatusBadge("3. Manifest: Waiting...");
            _badgeFss = CreateStatusBadge("4. FSS Sales CSV: Waiting...");

            _statusTable.Controls.Add(_badgeAcc, 0, 0);
            _statusTable.Controls.Add(_badgeOrd, 1, 0);
            _statusTable.Controls.Add(_badgeMan, 2, 0);
            _statusTable.Controls.Add(_badgeFss, 3, 0);

            _dropCard.Controls.AddRange(new Control[] {
                _lblDropTitle, _lblDropHint, _btnFilterOrders, _btnBrowseFiles, _btnClearFiles, _statusTable
            });

            // 3. Live Multi-Tab Preview Area (178px wide tabs so no title is ever cut off)
            _previewCard = CreateCardPanel(new Rectangle(18, 228, 1124, 444));

            _tabPreview = new TabControl
            {
                Location = new Point(14, 12),
                Size = new Size(1096, 418),
                Font = fontTab,
                ItemSize = new Size(178, 30),
                SizeMode = TabSizeMode.Fixed
            };

            _dgvCapture = CreateStyledGrid();
            _dgvSeuss = CreateStyledGrid();
            _dgvIris = CreateStyledGrid();
            _dgvPivot = CreateStyledGrid();
            _dgvOrderPhoto = CreateStyledGrid();
            _dgvOrderSales = CreateStyledGrid();

            AddTabWithGrid("Capture Report", _dgvCapture);
            AddTabWithGrid("Dr. Seuss Report", _dgvSeuss);
            AddTabWithGrid("Iris Sales (EOV)", _dgvIris);
            AddTabWithGrid("Accounting Pivot", _dgvPivot);
            AddTabWithGrid("OrderReport-Photo", _dgvOrderPhoto);
            AddTabWithGrid("OrderReport-Sales", _dgvOrderSales);

            _previewCard.Controls.Add(_tabPreview);

            // 4. Bottom Action Bar (Spacious, Zero Overlap, UseMnemonic=false so '&' is visible)
            _footerCard = CreateCardPanel(new Rectangle(18, 684, 1124, 62));

            _chkExportCsvs = new CheckBox
            {
                Text = "Export standalone OrderReport-Photo.csv & OrderReport-Sales.csv",
                Checked = true,
                Font = fontRegular,
                Location = new Point(16, 9),
                Size = new Size(495, 22),
                AutoSize = false,
                UseMnemonic = false
            };

            _lblFooterStatus = new Label
            {
                Text = "Ready — Drop your cruise files above to preview & generate.",
                Font = fontRegular,
                Location = new Point(16, 34),
                Size = new Size(390, 20),
                AutoEllipsis = true,
                UseMnemonic = false
            };

            _lblCreatedBy = new Label
            {
                Text = "Created by Gil",
                Font = fontItalic,
                Location = new Point(414, 34),
                Size = new Size(100, 20),
                TextAlign = ContentAlignment.MiddleLeft
            };

            _btnCopyTab = CreatePillButton("Copy Active Tab", new Point(526, 13), new Size(148, 36), false);
            _btnCopyTab.Click += BtnCopyTab_Click;

            _btnOpenFolder = CreatePillButton("Open Output Folder", new Point(684, 13), new Size(168, 36), false);
            _btnOpenFolder.Enabled = false;
            _btnOpenFolder.Click += (s, e) =>
            {
                if (!string.IsNullOrEmpty(_lastOutputFolder) && Directory.Exists(_lastOutputFolder))
                    Process.Start(new ProcessStartInfo { FileName = _lastOutputFolder, UseShellExecute = true });
            };

            _btnGenerateAll = CreatePillButton("Generate Excel + PDF + JPG", new Point(862, 13), new Size(246, 36), true);
            _btnGenerateAll.Click += BtnGenerateAll_Click;

            _footerCard.Controls.AddRange(new Control[] {
                _chkExportCsvs, _lblFooterStatus, _lblCreatedBy, _btnCopyTab, _btnOpenFolder, _btnGenerateAll
            });

            Controls.AddRange(new Control[] { _headerCard, _dropCard, _previewCard, _footerCard });
        }

        private void AddTabWithGrid(string title, DataGridView grid)
        {
            var page = new TabPage(title) { Padding = new Padding(6) };
            grid.Dock = DockStyle.Fill;
            page.Controls.Add(grid);
            _tabPreview.TabPages.Add(page);
        }

        private Panel CreateCardPanel(Rectangle bounds)
        {
            var panel = new Panel { Bounds = bounds, BackColor = Color.Transparent };
            panel.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                var rect = new Rectangle(0, 0, panel.Width - 1, panel.Height - 1);
                using var path = GetRoundedRectPath(rect, 12);
                var fill = _isDarkMode ? Color.FromArgb(28, 28, 30) : Color.White;
                var border = _isDarkMode ? Color.FromArgb(58, 58, 62) : Color.FromArgb(224, 224, 230);
                using var brush = new SolidBrush(fill);
                using var pen = new Pen(border, 1f);
                e.Graphics.FillPath(brush, path);
                e.Graphics.DrawPath(pen, path);
            };
            return panel;
        }

        private Button CreatePillButton(string text, Point location, Size size, bool isPrimary)
        {
            var btn = new Button
            {
                Text = text,
                Location = location,
                Size = size,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI Semibold", 12.5f, FontStyle.Bold, GraphicsUnit.Pixel),
                TextAlign = ContentAlignment.MiddleCenter,
                UseVisualStyleBackColor = false,
                UseMnemonic = false,
                Tag = isPrimary ? "primary" : "secondary"
            };
            btn.FlatAppearance.BorderSize = 1;
            return btn;
        }

        private Label CreateStatusBadge(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                Margin = new Padding(4, 3, 4, 3),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 6, 0),
                Font = new Font("Segoe UI Semibold", 12f, FontStyle.Bold, GraphicsUnit.Pixel),
                AutoEllipsis = true,
                UseMnemonic = false,
                Tag = "waiting"
            };
        }

        private DataGridView CreateStyledGrid()
        {
            var dgv = new DataGridView
            {
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 32,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                AllowUserToResizeRows = false,
                Font = new Font("Segoe UI", 12f, FontStyle.Regular, GraphicsUnit.Pixel)
            };
            dgv.RowTemplate.Height = 26;
            return dgv;
        }

        private void ApplyTheme()
        {
            BackColor = _isDarkMode ? Color.FromArgb(16, 16, 18) : Color.FromArgb(242, 242, 247);
            var textPrimary = _isDarkMode ? Color.FromArgb(245, 245, 247) : Color.FromArgb(28, 28, 30);
            var textSecondary = _isDarkMode ? Color.FromArgb(158, 158, 165) : Color.FromArgb(108, 108, 114);

            foreach (Control card in Controls)
            {
                card.Invalidate();
                foreach (Control c in card.Controls)
                {
                    if (c is Label lbl)
                    {
                        if (lbl == _lblDropHint || lbl == _lblFooterStatus || lbl == _lblCreatedBy)
                            lbl.ForeColor = textSecondary;
                        else
                            lbl.ForeColor = textPrimary;
                    }
                    else if (c is CheckBox chk)
                    {
                        chk.ForeColor = textPrimary;
                    }
                    else if (c is TextBox txt)
                    {
                        txt.BackColor = _isDarkMode ? Color.FromArgb(44, 44, 46) : Color.FromArgb(248, 248, 252);
                        txt.ForeColor = textPrimary;
                    }
                    else if (c is ComboBox cmb)
                    {
                        cmb.BackColor = _isDarkMode ? Color.FromArgb(44, 44, 46) : Color.White;
                        cmb.ForeColor = textPrimary;
                    }
                    else if (c is Button btn)
                    {
                        bool isPrimary = (btn.Tag as string) == "primary";
                        if (isPrimary)
                        {
                            btn.BackColor = Color.FromArgb(0, 122, 255);
                            btn.ForeColor = Color.White;
                            btn.FlatAppearance.BorderColor = Color.FromArgb(0, 105, 220);
                        }
                        else
                        {
                            btn.BackColor = _isDarkMode ? Color.FromArgb(44, 44, 46) : Color.FromArgb(238, 238, 244);
                            btn.ForeColor = textPrimary;
                            btn.FlatAppearance.BorderColor = _isDarkMode ? Color.FromArgb(72, 72, 76) : Color.FromArgb(208, 208, 216);
                        }
                    }
                }
            }

            UpdateBadgeColors();

            var grids = new[] { _dgvCapture, _dgvSeuss, _dgvIris, _dgvPivot, _dgvOrderPhoto, _dgvOrderSales };
            foreach (var g in grids)
            {
                g.BackgroundColor = _isDarkMode ? Color.FromArgb(28, 28, 30) : Color.White;
                g.GridColor = _isDarkMode ? Color.FromArgb(48, 48, 52) : Color.FromArgb(232, 232, 237);
                g.DefaultCellStyle.BackColor = _isDarkMode ? Color.FromArgb(28, 28, 30) : Color.White;
                g.DefaultCellStyle.ForeColor = textPrimary;
                g.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 122, 255);
                g.DefaultCellStyle.SelectionForeColor = Color.White;
                g.ColumnHeadersDefaultCellStyle.BackColor = _isDarkMode ? Color.FromArgb(44, 44, 46) : Color.FromArgb(242, 242, 247);
                g.ColumnHeadersDefaultCellStyle.ForeColor = textPrimary;
                g.ColumnHeadersDefaultCellStyle.Font = _gridHeaderFont;
            }

            foreach (TabPage tp in _tabPreview.TabPages)
            {
                tp.BackColor = _isDarkMode ? Color.FromArgb(28, 28, 30) : Color.White;
            }
        }

        private void UpdateBadgeColors()
        {
            var badges = new[] { _badgeAcc, _badgeOrd, _badgeMan, _badgeFss };
            foreach (var b in badges)
            {
                bool loaded = (b.Tag as string) == "loaded";
                if (loaded)
                {
                    b.BackColor = _isDarkMode ? Color.FromArgb(24, 68, 38) : Color.FromArgb(225, 246, 231);
                    b.ForeColor = _isDarkMode ? Color.FromArgb(110, 235, 140) : Color.FromArgb(20, 110, 46);
                }
                else
                {
                    b.BackColor = _isDarkMode ? Color.FromArgb(44, 44, 46) : Color.FromArgb(242, 242, 247);
                    b.ForeColor = _isDarkMode ? Color.FromArgb(160, 160, 165) : Color.FromArgb(110, 110, 115);
                }
            }
        }

        private static GraphicsPath GetRoundedRectPath(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            int d = radius * 2;
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        private void BtnCustomColor_Click(object sender, EventArgs e)
        {
            using var cd = new ColorDialog { Color = _activeMainColor, FullOpen = true };
            if (cd.ShowDialog(this) == DialogResult.OK)
            {
                _activeMainColor = cd.Color;
                int r = (int)(_activeMainColor.R * 0.25 + 255 * 0.75);
                int g = (int)(_activeMainColor.G * 0.25 + 255 * 0.75);
                int b = (int)(_activeMainColor.B * 0.25 + 255 * 0.75);
                _activeSubColor = Color.FromArgb(r, g, b);
                _pnlColorPreview.Invalidate();
            }
        }

        private void MainForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void MainForm_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data?.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
            {
                ProcessDroppedFiles(files);
            }
        }

        private void BtnBrowseFiles_Click(object sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Title = "Select Your Cruise Report Files (Accounting, Order, Manifest, FSS CSV)",
                Multiselect = true,
                Filter = "All Report Files (*.xls;*.xlsx;*.csv)|*.xls;*.xlsx;*.csv|All Files (*.*)|*.*"
            };
            if (ofd.ShowDialog(this) == DialogResult.OK)
            {
                ProcessDroppedFiles(ofd.FileNames);
            }
        }

        private void ClearAllData()
        {
            _accountingFilePath = _orderFilePath = _manifestFilePath = _fssFilePath = "";
            _accountingRows.Clear();
            _rawOrderRows.Clear();
            _orderRowsForDataTab.Clear();
            _orderRowsForPhotoSales.Clear();
            _manifestRows.Clear();
            _irisSummary = new IrisSummary();

            _badgeAcc.Text = "1. Accounting: Waiting...";
            _badgeAcc.Tag = "waiting";
            _badgeOrd.Text = "2. Order Report: Waiting...";
            _badgeOrd.Tag = "waiting";
            _badgeMan.Text = "3. Manifest: Waiting...";
            _badgeMan.Tag = "waiting";
            _badgeFss.Text = "4. FSS Sales CSV: Waiting...";
            _badgeFss.Tag = "waiting";
            UpdateBadgeColors();

            _dgvCapture.DataSource = null;
            _dgvSeuss.DataSource = null;
            _dgvIris.DataSource = null;
            _dgvPivot.DataSource = null;
            _dgvOrderPhoto.DataSource = null;
            _dgvOrderSales.DataSource = null;

            _lblFooterStatus.Text = "Cleared all loaded data.";
        }

        private void ProcessDroppedFiles(IEnumerable<string> paths)
        {
            var allFiles = new List<string>();
            foreach (var p in paths)
            {
                if (Directory.Exists(p))
                    allFiles.AddRange(Directory.GetFiles(p, "*.*", SearchOption.TopDirectoryOnly));
                else if (File.Exists(p))
                    allFiles.Add(p);
            }

            foreach (var file in allFiles)
            {
                try
                {
                    AutoDetectAndLoadFile(file);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"Error reading {Path.GetFileName(file)}:\n{ex.Message}", "File Parse Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            UpdateBadgeColors();
            RefreshPreviews();
        }

        private void AutoDetectAndLoadFile(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            string fileName = Path.GetFileName(filePath);

            var mVoy = Regex.Match(fileName, @"([A-Z]{2}\d{6})", RegexOptions.IgnoreCase);
            if (mVoy.Success)
                _txtVoyage.Text = mVoy.Groups[1].Value.ToUpperInvariant();

            if (ext == ".csv")
            {
                string head = File.ReadLines(filePath).Take(10).Aggregate("", (a, b) => a + "\n" + b);
                if (head.IndexOf("FSS POS ITEM SALES REPORT", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    head.IndexOf("IRIS PHOTO", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    head.IndexOf("PLU Number", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    ParseFssCsv(filePath);
                    _fssFilePath = filePath;
                    _badgeFss.Text = $"✓ FSS: {_irisSummary.TotalQty} Iris (${_irisSummary.TotalAmt:N2})";
                    _badgeFss.Tag = "loaded";
                }
                return;
            }

            string firstBytes = "";
            using (var sr = new StreamReader(filePath, Encoding.UTF8, true))
            {
                char[] buf = new char[8192];
                int n = sr.Read(buf, 0, buf.Length);
                firstBytes = new string(buf, 0, n);
            }

            if (firstBytes.IndexOf("<table", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (firstBytes.IndexOf("Account Report Statistics", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    ParseAccountingHtmlXls(filePath);
                    _accountingFilePath = filePath;
                    _badgeAcc.Text = $"✓ Accounting: {_accountingRows.Count} rows ({_accountingRows.Sum(r => r.Images):N0} imgs)";
                    _badgeAcc.Tag = "loaded";
                }
                else if (firstBytes.IndexOf("Order Report", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    ParseOrderHtmlXls(filePath);
                    _orderFilePath = filePath;
                    _badgeOrd.Text = $"✓ Orders: {_orderRowsForDataTab.Count:N0} Data / {_orderRowsForPhotoSales.Count:N0} Sales";
                    _badgeOrd.Tag = "loaded";
                }
                return;
            }

            if (ext == ".xlsx" || ext == ".xlsm")
            {
                using var wb = new XLWorkbook(filePath);
                var ws = wb.Worksheets.First();
                string a1 = ws.Cell(1, 1).GetString().Trim();
                string b1 = ws.Cell(1, 2).GetString().Trim();
                if (a1.Equals("CABIN_NUM", StringComparison.OrdinalIgnoreCase) ||
                    b1.Equals("ITEM_TAG", StringComparison.OrdinalIgnoreCase))
                {
                    ParseManifestXlsx(wb);
                    _manifestFilePath = filePath;
                    _badgeMan.Text = $"✓ Manifest: {_manifestRows.Count} Pre-Cruise pkgs";
                    _badgeMan.Tag = "loaded";
                }
            }
        }

        private static string NormalizeSpaces(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "";
            return Regex.Replace(input.Trim(), @"\s+", " ");
        }

        private string ClassifyCaptureType(string captureType)
        {
            string clean = NormalizeSpaces(captureType).ToUpperInvariant();
            if (PluCategoryMap.TryGetValue(clean, out string cat))
                return cat;

            string[] ccKeywords = { "GANGWAY", "RESTAURANT", "EMBARKATION", "DECK", "WHEEL", "BREAKFAST", "CLASSICS", "PARROTS", "PIRATE", "TOWEL", "INDIAN", "MARIACHI", "DANCER", "SKYRIDE", "SLIDES" };
            foreach (var kw in ccKeywords)
            {
                if (clean.Contains(kw)) return "CC";
            }
            return "PP";
        }

        private void ParseAccountingHtmlXls(string filePath)
        {
            _accountingRows.Clear();
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(filePath, Encoding.UTF8);

            var tables = doc.DocumentNode.SelectNodes("//table");
            if (tables == null || tables.Count == 0) return;

            HtmlNode dataTable = tables.LastOrDefault(t => (t.SelectNodes(".//tr")?.Count ?? 0) > 2) ?? tables.Last();
            var rows = dataTable.SelectNodes(".//tr");
            if (rows == null) return;

            int headerIdx = -1;
            var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < rows.Count; i++)
            {
                var cells = rows[i].SelectNodes("./td|./th");
                if (cells != null && cells.Count >= 8 && cells[0].InnerText.Trim().Equals("Employee ID", StringComparison.OrdinalIgnoreCase))
                {
                    headerIdx = i;
                    for (int c = 0; c < cells.Count; c++)
                    {
                        string hName = NormalizeSpaces(HtmlEntity.DeEntitize(cells[c].InnerText));
                        if (!string.IsNullOrEmpty(hName) && !colMap.ContainsKey(hName))
                            colMap[hName] = c;
                    }
                    break;
                }
            }

            if (headerIdx == -1) return;

            int idxEmp = colMap.TryGetValue("Employee ID", out int cEmp) ? cEmp : 0;
            int idxPhotog = colMap.TryGetValue("Photographer Name", out int cPhotog) ? cPhotog : 1;
            int idxCap = colMap.TryGetValue("Capture Type", out int cCap) ? cCap : 2;
            int idxPlu = colMap.TryGetValue("Legacy FSS PLU", out int cPlu) ? cPlu : 3;
            int idxImgs = colMap.TryGetValue("No.of Images", out int cImgs) ? cImgs : 4;
            int idxCore = colMap.TryGetValue("Quantity Sold-Core", out int cCore) ? cCore : 5;
            int idxUpsell = colMap.TryGetValue("Quantity Sold-Upsell", out int cUp) ? cUp : 7;
            int idxAmt = colMap.TryGetValue("Total Sale Amount", out int cAmt) ? cAmt : 10;
            int minCols = Math.Max(idxAmt, Math.Max(idxUpsell, idxCore)) + 1;

            for (int i = headerIdx + 1; i < rows.Count; i++)
            {
                var cells = rows[i].SelectNodes("./td|./th");
                if (cells == null || cells.Count < minCols) continue;

                string empId = HtmlEntity.DeEntitize(cells[idxEmp].InnerText).Trim();
                if (string.IsNullOrEmpty(empId)) continue;

                string photog = NormalizeSpaces(HtmlEntity.DeEntitize(cells[idxPhotog].InnerText));
                string capType = HtmlEntity.DeEntitize(cells[idxCap].InnerText).Trim();
                string plu = HtmlEntity.DeEntitize(cells[idxPlu].InnerText).Trim();

                int.TryParse(HtmlEntity.DeEntitize(cells[idxImgs].InnerText).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out int imgs);
                int.TryParse(HtmlEntity.DeEntitize(cells[idxCore].InnerText).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out int soldCore);
                int.TryParse(HtmlEntity.DeEntitize(cells[idxUpsell].InnerText).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out int soldUpsell);
                double.TryParse(HtmlEntity.DeEntitize(cells[idxAmt].InnerText).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out double totalAmt);

                _accountingRows.Add(new AccountingRow
                {
                    EmployeeId = empId,
                    PhotographerName = photog,
                    CaptureType = capType,
                    Plu = plu,
                    Images = imgs,
                    SoldCore = soldCore,
                    SoldUpsell = soldUpsell,
                    TotalAmount = Math.Round(totalAmt, 2),
                    Category = ClassifyCaptureType(capType)
                });
            }
        }

        private void ParseOrderHtmlXls(string filePath)
        {
            _rawOrderRows.Clear();
            _orderRowsForDataTab.Clear();
            _orderRowsForPhotoSales.Clear();

            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(filePath, Encoding.UTF8);

            var tables = doc.DocumentNode.SelectNodes("//table");
            if (tables == null || tables.Count == 0) return;

            HtmlNode dataTable = tables.LastOrDefault(t => (t.SelectNodes(".//tr")?.Count ?? 0) > 2) ?? tables.Last();
            var rows = dataTable.SelectNodes(".//tr");
            if (rows == null) return;

            int headerIdx = -1;
            var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < rows.Count; i++)
            {
                var cells = rows[i].SelectNodes("./td|./th");
                if (cells != null && cells.Count >= 15 && cells[0].InnerText.Trim().Equals("Order ID", StringComparison.OrdinalIgnoreCase))
                {
                    headerIdx = i;
                    for (int c = 0; c < cells.Count; c++)
                    {
                        string hName = NormalizeSpaces(HtmlEntity.DeEntitize(cells[c].InnerText));
                        if (!string.IsNullOrEmpty(hName) && !colMap.ContainsKey(hName))
                            colMap[hName] = c;
                    }
                    break;
                }
            }

            if (headerIdx == -1) return;

            int idxOrdId = colMap.TryGetValue("Order ID", out int cOrd) ? cOrd : 0;
            int idxCabin = colMap.TryGetValue("Cabin No", out int cCab) ? cCab : (colMap.TryGetValue("Cabin Number", out int cCab2) ? cCab2 : 2);
            int idxGuest = colMap.TryGetValue("Guest Name", out int cGst) ? cGst : 4;
            int idxFile = colMap.TryGetValue("File Name", out int cFil) ? cFil : 5;
            int idxEmp = colMap.TryGetValue("Employee ID", out int cEmp) ? cEmp : 6;
            int idxPhotog = colMap.TryGetValue("Photographer Name", out int cPho) ? cPho : 7;
            int idxCap = colMap.TryGetValue("Capture Type", out int cCap) ? cCap : 11;
            int idxMedia = colMap.TryGetValue("Media Type", out int cMed) ? cMed : 12;
            int idxStatus = colMap.TryGetValue("Order Status", out int cSts) ? cSts : 17;
            int idxVoyage = colMap.TryGetValue("Voyage Number", out int cVoy) ? cVoy : 18;
            int idxChannel = colMap.TryGetValue("ChannelID", out int cChn) ? cChn : (colMap.TryGetValue("Channel ID", out int cChn2) ? cChn2 : 21);
            int minCols = new[] { idxOrdId, idxCabin, idxGuest, idxFile, idxEmp, idxPhotog, idxCap, idxMedia, idxStatus, idxVoyage, idxChannel }.Max() + 1;

            for (int i = headerIdx + 1; i < rows.Count; i++)
            {
                var cells = rows[i].SelectNodes("./td|./th");
                if (cells == null || cells.Count < minCols) continue;

                string orderStatus = HtmlEntity.DeEntitize(cells[idxStatus].InnerText).Trim();

                var item = new OrderRow
                {
                    OrderId = HtmlEntity.DeEntitize(cells[idxOrdId].InnerText).Trim(),
                    CabinNo = HtmlEntity.DeEntitize(cells[idxCabin].InnerText).Trim(),
                    GuestName = NormalizeSpaces(HtmlEntity.DeEntitize(cells[idxGuest].InnerText)),
                    FileName = HtmlEntity.DeEntitize(cells[idxFile].InnerText).Trim(),
                    EmployeeId = HtmlEntity.DeEntitize(cells[idxEmp].InnerText).Trim(),
                    PhotographerName = NormalizeSpaces(HtmlEntity.DeEntitize(cells[idxPhotog].InnerText)),
                    CaptureType = HtmlEntity.DeEntitize(cells[idxCap].InnerText).Trim(),
                    MediaType = HtmlEntity.DeEntitize(cells[idxMedia].InnerText).Trim(),
                    OrderStatus = orderStatus,
                    VoyageNumber = HtmlEntity.DeEntitize(cells[idxVoyage].InnerText).Trim(),
                    ChannelId = HtmlEntity.DeEntitize(cells[idxChannel].InnerText).Trim()
                };

                _rawOrderRows.Add(item);
            }

            ApplyOrderFilters();
            ShowOrderFilterChecklistDialog();
        }

        private void ApplyOrderFilters()
        {
            _orderRowsForDataTab.Clear();
            _orderRowsForPhotoSales.Clear();

            foreach (var item in _rawOrderRows)
            {
                if (_excludedOrderStatuses.Contains(item.OrderStatus))
                    continue;

                _orderRowsForDataTab.Add(item);

                if (!_excludedMediaTypes.Contains(item.MediaType))
                {
                    _orderRowsForPhotoSales.Add(item);
                }
            }

            if (!string.IsNullOrEmpty(_orderFilePath))
            {
                _badgeOrd.Text = $"✓ Orders: {_orderRowsForDataTab.Count:N0} Data / {_orderRowsForPhotoSales.Count:N0} Sales";
                _badgeOrd.Tag = "loaded";
            }
        }

        private void ShowOrderFilterChecklistDialog()
        {
            if (_rawOrderRows.Count == 0)
            {
                MessageBox.Show(this, "Drop or browse an OrderReport (.xls) file first to configure Media Type and Order Status filters.", "Order Filters", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var dlg = new Form
            {
                AutoScaleMode = AutoScaleMode.None,
                Text = "Filter / Exclude Media Types & Order Statuses",
                ClientSize = new Size(520, 420),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = _isDarkMode ? Color.FromArgb(28, 28, 30) : Color.FromArgb(242, 242, 247),
                ForeColor = _isDarkMode ? Color.FromArgb(245, 245, 247) : Color.FromArgb(28, 28, 30)
            };

            using var fontHeader = new Font("Segoe UI Semibold", 13f, FontStyle.Bold, GraphicsUnit.Pixel);
            using var fontSub = new Font("Segoe UI", 12f, FontStyle.Regular, GraphicsUnit.Pixel);

            var lblInstruction = new Label
            {
                Text = "Check items below to EXCLUDE them from OrderReport-Photo & OrderReport-Sales:",
                Font = fontHeader,
                Location = new Point(16, 12),
                Size = new Size(488, 22),
                UseMnemonic = false
            };

            var lblMedia = new Label
            {
                Text = "Exclude Media Types (Photo & Sales):",
                Font = fontHeader,
                Location = new Point(16, 42),
                Size = new Size(236, 20),
                UseMnemonic = false
            };

            var lblStatus = new Label
            {
                Text = "Exclude Order Status (All Order Tabs):",
                Font = fontHeader,
                Location = new Point(268, 42),
                Size = new Size(236, 20),
                UseMnemonic = false
            };

            var clbMedia = new CheckedListBox
            {
                Location = new Point(16, 66),
                Size = new Size(236, 290),
                CheckOnClick = true,
                Font = fontSub,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = _isDarkMode ? Color.FromArgb(44, 44, 46) : Color.White,
                ForeColor = _isDarkMode ? Color.FromArgb(245, 245, 247) : Color.FromArgb(28, 28, 30)
            };

            var clbStatus = new CheckedListBox
            {
                Location = new Point(268, 66),
                Size = new Size(236, 290),
                CheckOnClick = true,
                Font = fontSub,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = _isDarkMode ? Color.FromArgb(44, 44, 46) : Color.White,
                ForeColor = _isDarkMode ? Color.FromArgb(245, 245, 247) : Color.FromArgb(28, 28, 30)
            };

            var defaultMediaCodes = new[] { "OB02", "OB03", "OB04", "OB05" };
            var distinctMedia = _rawOrderRows
                .Select(r => r.MediaType)
                .Where(m => !string.IsNullOrWhiteSpace(m))
                .Union(defaultMediaCodes, StringComparer.OrdinalIgnoreCase)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(m => m, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var m in distinctMedia)
            {
                bool isExcluded = _excludedMediaTypes.Contains(m) ||
                                  (m.Equals("OB02", StringComparison.OrdinalIgnoreCase) && _excludedMediaTypes.Contains("OB2")) ||
                                  (m.Equals("OB03", StringComparison.OrdinalIgnoreCase) && _excludedMediaTypes.Contains("OB3")) ||
                                  (m.Equals("OB04", StringComparison.OrdinalIgnoreCase) && _excludedMediaTypes.Contains("OB4")) ||
                                  (m.Equals("OB05", StringComparison.OrdinalIgnoreCase) && _excludedMediaTypes.Contains("OB5"));
                clbMedia.Items.Add(m, isExcluded);
            }

            var defaultStatuses = new[] { "Refund", "InCart" };
            var distinctStatuses = _rawOrderRows
                .Select(r => r.OrderStatus)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Union(defaultStatuses, StringComparer.OrdinalIgnoreCase)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var st in distinctStatuses)
            {
                bool isExcluded = _excludedOrderStatuses.Contains(st);
                clbStatus.Items.Add(st, isExcluded);
            }

            var btnSelectDefault = CreatePillButton("Reset Defaults", new Point(16, 368), new Size(130, 36), false);
            btnSelectDefault.Click += (s, e) =>
            {
                var defM = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "OB02", "OB03", "OB04", "OB05", "OB2", "OB3", "OB4", "OB5" };
                for (int i = 0; i < clbMedia.Items.Count; i++)
                    clbMedia.SetItemChecked(i, defM.Contains(clbMedia.Items[i].ToString()));

                var defS = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Refund", "InCart" };
                for (int i = 0; i < clbStatus.Items.Count; i++)
                    clbStatus.SetItemChecked(i, defS.Contains(clbStatus.Items[i].ToString()));
            };

            var btnApply = CreatePillButton("Apply Filters", new Point(354, 368), new Size(150, 36), true);
            btnApply.BackColor = Color.FromArgb(0, 122, 255);
            btnApply.ForeColor = Color.White;
            btnApply.Click += (s, e) =>
            {
                _excludedMediaTypes.Clear();
                foreach (var item in clbMedia.CheckedItems)
                {
                    string val = item.ToString();
                    _excludedMediaTypes.Add(val);
                    if (val.Equals("OB02", StringComparison.OrdinalIgnoreCase)) _excludedMediaTypes.Add("OB2");
                    if (val.Equals("OB03", StringComparison.OrdinalIgnoreCase)) _excludedMediaTypes.Add("OB3");
                    if (val.Equals("OB04", StringComparison.OrdinalIgnoreCase)) _excludedMediaTypes.Add("OB4");
                    if (val.Equals("OB05", StringComparison.OrdinalIgnoreCase)) _excludedMediaTypes.Add("OB5");
                }

                _excludedOrderStatuses.Clear();
                foreach (var item in clbStatus.CheckedItems)
                {
                    _excludedOrderStatuses.Add(item.ToString());
                }

                ApplyOrderFilters();
                RefreshPreviews();
                dlg.DialogResult = DialogResult.OK;
                dlg.Close();
            };

            dlg.Controls.AddRange(new Control[] { lblInstruction, lblMedia, lblStatus, clbMedia, clbStatus, btnSelectDefault, btnApply });
            dlg.ShowDialog(this);
        }

        private void ParseManifestXlsx(XLWorkbook wb)
        {
            _manifestRows.Clear();
            var ws = wb.Worksheets.First();
            int lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

            for (int r = 2; r <= lastRow; r++)
            {
                string itemTag = ws.Cell(r, 2).GetString().Trim();
                string itemDesc = ws.Cell(r, 4).GetString().Trim();

                if ((itemTag.Equals("PHOTO", StringComparison.OrdinalIgnoreCase) ||
                     itemTag.Equals("PIXELSPKG", StringComparison.OrdinalIgnoreCase)) &&
                    itemDesc.IndexOf("ProPortraits Package", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string cabinRaw = ws.Cell(r, 1).GetString().Trim();
                    if (double.TryParse(cabinRaw, NumberStyles.Any, CultureInfo.InvariantCulture, out double cabNum))
                        cabinRaw = ((int)cabNum).ToString();

                    int.TryParse(ws.Cell(r, 3).GetString().Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out int qty);
                    if (qty == 0) qty = 1;

                    string guestName = ws.Cell(r, 6).GetString().Trim();

                    _manifestRows.Add(new ManifestRow
                    {
                        CabinNum = cabinRaw,
                        Quantity = qty,
                        ItemDescription = itemDesc,
                        GuestName = guestName
                    });
                }
            }
        }

        private static List<string> SplitCsvLine(string line)
        {
            var result = new List<string>();
            if (string.IsNullOrEmpty(line)) return result;

            var sb = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(sb.ToString());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }
            result.Add(sb.ToString());
            return result;
        }

        private static string SanitizeSpreadsheetText(string val)
        {
            if (string.IsNullOrEmpty(val)) return "";
            char first = val[0];
            if (first == '=' || first == '+' || first == '-' || first == '@')
                return "'" + val;
            return val;
        }

        private static string EscapeCsvField(string val)
        {
            if (string.IsNullOrEmpty(val)) return "";
            string clean = SanitizeSpreadsheetText(val);
            if (clean.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0)
                return "\"" + clean.Replace("\"", "\"\"") + "\"";
            return clean;
        }

        private void ParseFssCsv(string filePath)
        {
            _irisSummary = new IrisSummary();
            var lines = File.ReadAllLines(filePath);
            bool inIrisSection = false;

            foreach (var rawLine in lines)
            {
                string line = rawLine.Trim();
                if (string.IsNullOrEmpty(line)) continue;

                if (line.Equals("IRIS PHOTO", StringComparison.OrdinalIgnoreCase))
                {
                    inIrisSection = true;
                    continue;
                }

                if (inIrisSection && line.StartsWith("Sub Total For IRIS PHOTO", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }

                if (inIrisSection)
                {
                    var parts = SplitCsvLine(line);
                    if (parts.Length < 6) continue;

                    string plu = parts[1].Trim();
                    string desc = parts[2].Trim().ToUpperInvariant();
                    int.TryParse(parts[3].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out int qty);
                    string amtStr = parts[5].Trim().Replace("$", "").Replace(",", "");
                    double.TryParse(amtStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double amt);

                    if (plu == "111001504" || desc.Contains("PREMIUM PACK"))
                    {
                        _irisSummary.PremiumPkgQty += qty;
                        _irisSummary.PremiumPkgAmt += amt;
                    }
                    else if (plu == "111001499" || desc.Contains("ADD ON"))
                    {
                        _irisSummary.AddOnQty += qty;
                        _irisSummary.AddOnAmt += amt;
                    }
                    else if (plu == "111001503" || plu == "111001616" || plu == "111001665" || plu == "111001669" || desc.Contains("DIGITAL FILE"))
                    {
                        _irisSummary.DigitalUsbQty += qty;
                        _irisSummary.DigitalUsbAmt += amt;
                    }
                    else if (desc.Contains("CANVAS"))
                    {
                        _irisSummary.CanvasQty += qty;
                        _irisSummary.CanvasAmt += amt;
                    }
                    else if (desc.Contains("10X10"))
                    {
                        _irisSummary.Print10x10Qty += qty;
                        _irisSummary.Print10x10Amt += amt;
                    }
                    else if (desc.Contains("8X10"))
                    {
                        _irisSummary.Print8x10Qty += qty;
                        _irisSummary.Print8x10Amt += amt;
                    }
                    else if (desc.Contains("10X13"))
                    {
                        _irisSummary.Print10x13Qty += qty;
                        _irisSummary.Print10x13Amt += amt;
                    }
                }
            }
        }

        private void RefreshPreviews()
        {
            string voyage = _txtVoyage.Text.Trim();

            if (_accountingRows.Count > 0)
            {
                var dtCap = new DataTable();
                dtCap.Columns.Add("CATEGORY");
                dtCap.Columns.Add("TOTAL CAPTURES", typeof(int));
                dtCap.Columns.Add("TOTAL SOLD", typeof(int));
                dtCap.Columns.Add("TOTAL AMOUNT", typeof(double));
                dtCap.Columns.Add("AVERAGE");

                var ccRows = _accountingRows.Where(r => r.Category == "CC").ToList();
                var ppRows = _accountingRows.Where(r => r.Category == "PP").ToList();

                int ccImg = ccRows.Sum(r => r.Images);
                int ccSold = ccRows.Sum(r => r.SoldCore);
                double ccAmt = Math.Round(ccRows.Sum(r => r.TotalAmount), 2);
                double ccAvg = ccImg > 0 ? (double)ccSold / ccImg : 0;

                int ppImg = ppRows.Sum(r => r.Images);
                int ppSold = ppRows.Sum(r => r.SoldCore);
                double ppAmt = Math.Round(ppRows.Sum(r => r.TotalAmount), 2);
                double ppAvg = ppImg > 0 ? (double)ppSold / ppImg : 0;

                int totImg = ccImg + ppImg;
                int totSold = ccSold + ppSold;
                double totAmt = Math.Round(ccAmt + ppAmt, 2);
                double totAvg = totImg > 0 ? (double)totSold / totImg : 0;

                dtCap.Rows.Add("CC", ccImg, ccSold, ccAmt, $"{ccAvg:P2}");
                dtCap.Rows.Add("PP", ppImg, ppSold, ppAmt, $"{ppAvg:P2}");
                dtCap.Rows.Add("TOTAL", totImg, totSold, totAmt, $"{totAvg:P2}");
                _dgvCapture.DataSource = dtCap;

                var dtSeuss = new DataTable();
                dtSeuss.Columns.Add("Voyage");
                dtSeuss.Columns.Add("Capture Type");
                dtSeuss.Columns.Add("No.of Images", typeof(int));
                dtSeuss.Columns.Add("Quantity Sold-Core", typeof(int));
                dtSeuss.Columns.Add("Total Sale Amount", typeof(double));

                var seussGroups = _accountingRows
                    .Where(r => r.CaptureType.IndexOf("SEUSS", StringComparison.OrdinalIgnoreCase) >= 0)
                    .GroupBy(r => r.CaptureType)
                    .OrderBy(g => g.Key)
                    .ToList();

                bool first = true;
                foreach (var g in seussGroups)
                {
                    dtSeuss.Rows.Add(
                        first ? voyage : "",
                        g.Key,
                        g.Sum(x => x.Images),
                        g.Sum(x => x.SoldCore),
                        Math.Round(g.Sum(x => x.TotalAmount), 2)
                    );
                    first = false;
                }
                dtSeuss.Rows.Add(
                    "TOTAL",
                    "",
                    seussGroups.Sum(g => g.Sum(x => x.Images)),
                    seussGroups.Sum(g => g.Sum(x => x.SoldCore)),
                    Math.Round(seussGroups.Sum(g => g.Sum(x => x.TotalAmount)), 2)
                );
                _dgvSeuss.DataSource = dtSeuss;

                var dtPivot = new DataTable();
                dtPivot.Columns.Add(voyage);
                dtPivot.Columns.Add("Sum of No.of Images", typeof(int));
                dtPivot.Columns.Add("Sum of Quantity Sold-Core", typeof(int));
                dtPivot.Columns.Add("Sum of Quantity Sold-Upsell", typeof(int));
                dtPivot.Columns.Add("Sum of Total Sale Amount", typeof(double));

                foreach (var pg in _accountingRows.GroupBy(r => r.PhotographerName).OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
                {
                    dtPivot.Rows.Add(
                        pg.Key,
                        pg.Sum(x => x.Images),
                        pg.Sum(x => x.SoldCore),
                        pg.Sum(x => x.SoldUpsell),
                        Math.Round(pg.Sum(x => x.TotalAmount), 2)
                    );
                    foreach (var cg in pg.GroupBy(x => x.CaptureType).OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
                    {
                        dtPivot.Rows.Add(
                            "    " + cg.Key,
                            cg.Sum(x => x.Images),
                            cg.Sum(x => x.SoldCore),
                            cg.Sum(x => x.SoldUpsell),
                            Math.Round(cg.Sum(x => x.TotalAmount), 2)
                        );
                    }
                }
                dtPivot.Rows.Add(
                    "Grand Total",
                    _accountingRows.Sum(x => x.Images),
                    _accountingRows.Sum(x => x.SoldCore),
                    _accountingRows.Sum(x => x.SoldUpsell),
                    Math.Round(_accountingRows.Sum(x => x.TotalAmount), 2)
                );
                _dgvPivot.DataSource = dtPivot;
            }

            var dtIris = new DataTable();
            dtIris.Columns.Add("VOYAGE");
            dtIris.Columns.Add("PREMIUM PKG", typeof(double));
            dtIris.Columns.Add("CANVAS", typeof(double));
            dtIris.Columns.Add("DIGITAL ON USB", typeof(double));
            dtIris.Columns.Add("10X10 PRINT", typeof(double));
            dtIris.Columns.Add("8X10 PRINT", typeof(double));
            dtIris.Columns.Add("10X13 PRINT", typeof(double));
            dtIris.Columns.Add("$19.99 ADD ON", typeof(double));
            dtIris.Columns.Add("TOTAL", typeof(double));

            dtIris.Rows.Add(
                voyage,
                _irisSummary.PremiumPkgQty,
                _irisSummary.CanvasQty,
                _irisSummary.DigitalUsbQty,
                _irisSummary.Print10x10Qty,
                _irisSummary.Print8x10Qty,
                _irisSummary.Print10x13Qty,
                _irisSummary.AddOnQty,
                _irisSummary.TotalQty
            );
            dtIris.Rows.Add(
                "",
                Math.Round(_irisSummary.PremiumPkgAmt, 2),
                Math.Round(_irisSummary.CanvasAmt, 2),
                Math.Round(_irisSummary.DigitalUsbAmt, 2),
                Math.Round(_irisSummary.Print10x10Amt, 2),
                Math.Round(_irisSummary.Print8x10Amt, 2),
                Math.Round(_irisSummary.Print10x13Amt, 2),
                Math.Round(_irisSummary.AddOnAmt, 2),
                _irisSummary.TotalAmt
            );
            _dgvIris.DataSource = dtIris;

            if (_rawOrderRows.Count > 0)
            {
                var dtPhoto = new DataTable();
                dtPhoto.Columns.Add("File Name");
                dtPhoto.Columns.Add("Photographer Name");

                var dtSales = new DataTable();
                dtSales.Columns.Add("Cabin No");
                dtSales.Columns.Add("Guest Name");
                dtSales.Columns.Add("File Name");
                dtSales.Columns.Add("Media Type");

                foreach (var r in _orderRowsForPhotoSales)
                {
                    dtPhoto.Rows.Add(r.FileName, r.PhotographerName);
                    dtSales.Rows.Add(r.CabinNo, r.GuestName, r.FileName, r.MediaType);
                }

                _dgvOrderPhoto.DataSource = dtPhoto;
                _dgvOrderSales.DataSource = dtSales;
            }
        }

        private void BtnCopyTab_Click(object sender, EventArgs e)
        {
            if (_tabPreview.SelectedTab?.Controls.OfType<DataGridView>().FirstOrDefault() is DataGridView dgv && dgv.Rows.Count > 0)
            {
                dgv.SelectAll();
                DataObject dataObj = dgv.GetClipboardContent();
                if (dataObj != null)
                {
                    Clipboard.SetDataObject(dataObj);
                    _lblFooterStatus.Text = $"Copied {dgv.Rows.Count} rows from '{_tabPreview.SelectedTab.Text}' to clipboard!";
                }
                dgv.ClearSelection();
            }
        }

        private void BtnGenerateAll_Click(object sender, EventArgs e)
        {
            if (_accountingRows.Count == 0 && _orderRowsForDataTab.Count == 0 && _manifestRows.Count == 0 && _irisSummary.TotalQty == 0)
            {
                MessageBox.Show(this, "Please drag & drop at least one report file before generating.", "No Data Loaded", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string voyage = string.IsNullOrWhiteSpace(_txtVoyage.Text) ? "PO102426" : _txtVoyage.Text.Trim();

            using var fbd = new FolderBrowserDialog
            {
                Description = "Select Output Folder for Generated Excel, PDF, JPG, and CSV Reports",
                UseDescriptionForTitle = true
            };

            string defaultDir = !string.IsNullOrEmpty(_accountingFilePath) ? Path.GetDirectoryName(_accountingFilePath)
                              : !string.IsNullOrEmpty(_orderFilePath) ? Path.GetDirectoryName(_orderFilePath)
                              : Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            fbd.SelectedPath = defaultDir;

            if (fbd.ShowDialog(this) != DialogResult.OK) return;

            string outDir = fbd.SelectedPath;
            _lastOutputFolder = outDir;

            try
            {
                Cursor = Cursors.WaitCursor;
                _lblFooterStatus.Text = "Generating multi-tab Excel, Accounting PDF, Sales.jpg, and CSVs...";
                Application.DoEvents();

                string excelPath = Path.Combine(outDir, $"{voyage}_Master_Generated.xlsx");
                GenerateMasterExcel(excelPath, voyage);

                string jpgPath = Path.Combine(outDir, "Sales.jpg");
                string pdfPath = Path.Combine(outDir, "AccountReportStatistics.pdf");
                int jpgHeightIn = 24;
                if (_accountingRows.Count > 0)
                {
                    var pivotLines = BuildPivotRenderLines(voyage);
                    jpgHeightIn = GetDynamicJpgHeightInches(pivotLines.Count);
                    GenerateSalesJpg(jpgPath, voyage, pivotLines);
                    GenerateAccountingPdf(pdfPath, voyage, pivotLines);
                }

                if (_chkExportCsvs.Checked && _orderRowsForPhotoSales.Count > 0)
                {
                    ExportOrderCsvs(outDir);
                }

                _btnOpenFolder.Enabled = true;
                _lblFooterStatus.Text = $"Success! Saved all reports to {outDir}";
                MessageBox.Show(this,
                    $"All reports generated successfully in:\n{outDir}\n\n" +
                    $"• {Path.GetFileName(excelPath)} (7 Tabs)\n" +
                    $"• AccountReportStatistics.pdf (4-Page US Letter)\n" +
                    $"• Sales.jpg (8x{jpgHeightIn}\" @ 300 DPI)\n" +
                    (_chkExportCsvs.Checked ? "• OrderReport-Photo.csv & OrderReport-Sales.csv" : ""),
                    "Generation Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Error generating files:\n{ex.Message}", "Generation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void GenerateMasterExcel(string filePath, string voyage)
        {
            using var wb = new XLWorkbook();
            var xlMainColor = XLColor.FromColor(_activeMainColor);
            var xlSubColor = XLColor.FromColor(_activeSubColor);

            // TAB 1: DATA_TAB_PASTE
            var wsData = wb.Worksheets.Add("DATA_TAB_PASTE");
            wsData.Cell(1, 1).Value = "Account Report Statistics";
            wsData.Cell(1, 9).Value = voyage;
            wsData.Cell(1, 11).Value = "Order Report";
            wsData.Cell(1, 15).Value = "PreCruise Package";
            wsData.Cell(2, 1).Value = DateTime.Now.ToString("MMMM", CultureInfo.InvariantCulture).ToUpperInvariant();

            string[] headersAtoI = { "Employee ID", "Photographer Name", "Capture Type", "Legacy FSS PLU", "No.of Images", "Quantity Sold-Core", "Quantity Sold-Upsell", "Total Sale Amount", "Category" };
            for (int i = 0; i < headersAtoI.Length; i++)
                wsData.Cell(4, i + 1).Value = headersAtoI[i];

            wsData.Cell(4, 11).Value = "Capture Type";
            wsData.Cell(4, 12).Value = "Media Type";
            wsData.Cell(4, 13).Value = "ChannelID";

            wsData.Cell(4, 15).Value = "CABIN_NUM";
            wsData.Cell(4, 16).Value = "QUANTITY";
            wsData.Cell(4, 17).Value = "ITEM_DESCRIPTION";
            wsData.Cell(4, 18).Value = "GUEST_NAME1";
            wsData.Cell(4, 19).Value = "AMOUNT";

            var headerRange = wsData.Range(4, 1, 4, 19);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = xlMainColor;

            for (int i = 0; i < _accountingRows.Count; i++)
            {
                int r = i + 5;
                var row = _accountingRows[i];
                if (int.TryParse(row.EmployeeId, out int empNum)) wsData.Cell(r, 1).Value = empNum;
                else wsData.Cell(r, 1).Value = row.EmployeeId;
                wsData.Cell(r, 2).Value = SanitizeSpreadsheetText(row.PhotographerName);
                wsData.Cell(r, 3).Value = row.CaptureType;
                if (long.TryParse(row.Plu, out long pluNum)) wsData.Cell(r, 4).Value = pluNum;
                else wsData.Cell(r, 4).Value = row.Plu;
                wsData.Cell(r, 5).Value = row.Images;
                wsData.Cell(r, 6).Value = row.SoldCore;
                wsData.Cell(r, 7).Value = row.SoldUpsell;
                wsData.Cell(r, 8).Value = row.TotalAmount;
                wsData.Cell(r, 9).Value = row.Category;
            }

            for (int i = 0; i < _orderRowsForDataTab.Count; i++)
            {
                int r = i + 5;
                var ord = _orderRowsForDataTab[i];
                wsData.Cell(r, 11).Value = ord.CaptureType;
                wsData.Cell(r, 12).Value = ord.MediaType;
                wsData.Cell(r, 13).Value = ord.ChannelId;
            }

            for (int i = 0; i < _manifestRows.Count; i++)
            {
                int r = i + 5;
                var man = _manifestRows[i];
                if (int.TryParse(man.CabinNum, out int cab)) wsData.Cell(r, 15).Value = cab;
                else wsData.Cell(r, 15).Value = man.CabinNum;
                wsData.Cell(r, 16).Value = man.Quantity;
                wsData.Cell(r, 17).Value = man.ItemDescription;
                wsData.Cell(r, 18).Value = SanitizeSpreadsheetText(man.GuestName);
            }

            // TAB 2: CAPTURE_REPORT
            var wsCap = wb.Worksheets.Add("CAPTURE_REPORT");
            wsCap.Cell(2, 3).Value = "TOTAL CAPTURES";
            wsCap.Cell(2, 4).Value = "TOTAL SOLD";
            wsCap.Cell(2, 5).Value = "TOTAL AMOUNT";
            wsCap.Cell(2, 6).Value = "AVERAGE";
            wsCap.Range(2, 2, 2, 6).Style.Font.Bold = true;
            wsCap.Range(2, 2, 2, 6).Style.Fill.BackgroundColor = xlMainColor;

            var ccList = _accountingRows.Where(x => x.Category == "CC").ToList();
            var ppList = _accountingRows.Where(x => x.Category == "PP").ToList();

            wsCap.Cell(3, 2).Value = "CC";
            wsCap.Cell(3, 3).Value = ccList.Sum(x => x.Images);
            wsCap.Cell(3, 4).Value = ccList.Sum(x => x.SoldCore);
            wsCap.Cell(3, 5).Value = Math.Round(ccList.Sum(x => x.TotalAmount), 2);
            wsCap.Cell(3, 6).FormulaA1 = "=IFERROR(D3/C3, 0)";
            wsCap.Cell(3, 6).Style.NumberFormat.Format = "0.00%";

            wsCap.Cell(4, 2).Value = "PP";
            wsCap.Cell(4, 3).Value = ppList.Sum(x => x.Images);
            wsCap.Cell(4, 4).Value = ppList.Sum(x => x.SoldCore);
            wsCap.Cell(4, 5).Value = Math.Round(ppList.Sum(x => x.TotalAmount), 2);
            wsCap.Cell(4, 6).FormulaA1 = "=IFERROR(D4/C4, 0)";
            wsCap.Cell(4, 6).Style.NumberFormat.Format = "0.00%";

            wsCap.Cell(5, 2).Value = "TOTAL";
            wsCap.Cell(5, 3).FormulaA1 = "=SUM(C3:C4)";
            wsCap.Cell(5, 4).FormulaA1 = "=SUM(D3:D4)";
            wsCap.Cell(5, 5).FormulaA1 = "=SUM(E3:E4)";
            wsCap.Cell(5, 6).FormulaA1 = "=IFERROR(D5/C5, 0)";
            wsCap.Cell(5, 6).Style.NumberFormat.Format = "0.00%";
            wsCap.Range(5, 2, 5, 6).Style.Font.Bold = true;
            wsCap.Range(5, 2, 5, 6).Style.Fill.BackgroundColor = xlSubColor;
            wsCap.Range(2, 2, 5, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            wsCap.Range(2, 2, 5, 6).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            wsCap.Columns().AdjustToContents();

            // TAB 3: DR_SEUSS_REPORT
            var wsSeuss = wb.Worksheets.Add("DR_SEUSS_REPORT");
            wsSeuss.Cell(1, 1).Value = "PANORAMA DR SEUSS REPORT";
            wsSeuss.Cell(1, 1).Style.Font.Bold = true;
            wsSeuss.Cell(1, 1).Style.Font.FontSize = 13;

            wsSeuss.Cell(3, 1).Value = "Voyage";
            wsSeuss.Cell(3, 2).Value = "Capture Type";
            wsSeuss.Cell(3, 3).Value = "No.of Images";
            wsSeuss.Cell(3, 4).Value = "Quantity Sold-Core";
            wsSeuss.Cell(3, 5).Value = "Total Sale Amount";
            wsSeuss.Range(3, 1, 3, 5).Style.Font.Bold = true;
            wsSeuss.Range(3, 1, 3, 5).Style.Fill.BackgroundColor = xlMainColor;

            var seussGroups = _accountingRows
                .Where(r => r.CaptureType.IndexOf("SEUSS", StringComparison.OrdinalIgnoreCase) >= 0)
                .GroupBy(r => r.CaptureType)
                .OrderBy(g => g.Key)
                .ToList();

            int sr = 4;
            foreach (var g in seussGroups)
            {
                if (sr == 4) wsSeuss.Cell(sr, 1).Value = voyage;
                wsSeuss.Cell(sr, 2).Value = g.Key;
                wsSeuss.Cell(sr, 3).Value = g.Sum(x => x.Images);
                wsSeuss.Cell(sr, 4).Value = g.Sum(x => x.SoldCore);
                wsSeuss.Cell(sr, 5).Value = Math.Round(g.Sum(x => x.TotalAmount), 2);
                sr++;
            }
            int totalRowSeuss = Math.Max(sr + 1, 8);
            wsSeuss.Cell(totalRowSeuss, 1).Value = "TOTAL";
            wsSeuss.Cell(totalRowSeuss, 3).FormulaA1 = $"=SUM(C4:C{totalRowSeuss - 1})";
            wsSeuss.Cell(totalRowSeuss, 4).FormulaA1 = $"=SUM(D4:D{totalRowSeuss - 1})";
            wsSeuss.Cell(totalRowSeuss, 5).FormulaA1 = $"=SUM(E4:E{totalRowSeuss - 1})";
            wsSeuss.Range(totalRowSeuss, 1, totalRowSeuss, 5).Style.Font.Bold = true;
            wsSeuss.Range(totalRowSeuss, 1, totalRowSeuss, 5).Style.Fill.BackgroundColor = xlSubColor;
            wsSeuss.Range(3, 1, totalRowSeuss, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            wsSeuss.Range(3, 1, totalRowSeuss, 5).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            wsSeuss.Columns().AdjustToContents();

            // TAB 4: IRIS_SALES
            var wsIris = wb.Worksheets.Add("IRIS_SALES");
            wsIris.Cell(1, 1).Value = "IRIS SALES (Copy C4:J5 directly to EOV Tab O66:V67)";
            wsIris.Cell(1, 1).Style.Font.Bold = true;

            wsIris.Cell(3, 1).Value = "VOYAGE";
            wsIris.Cell(3, 3).Value = "PREMIUM PKG";
            wsIris.Cell(3, 4).Value = "CANVAS";
            wsIris.Cell(3, 5).Value = "DIGITAL ON USB";
            wsIris.Cell(3, 6).Value = "10X10 PRINT";
            wsIris.Cell(3, 7).Value = "8X10 PRINT";
            wsIris.Cell(3, 8).Value = "10X13 PRINT";
            wsIris.Cell(3, 9).Value = "$19.99 ADD ON";
            wsIris.Cell(3, 10).Value = "TOTAL";
            wsIris.Range(3, 1, 3, 10).Style.Font.Bold = true;
            wsIris.Range(3, 1, 3, 10).Style.Fill.BackgroundColor = xlMainColor;

            wsIris.Cell(4, 1).Value = voyage;
            wsIris.Cell(4, 3).Value = _irisSummary.PremiumPkgQty;
            wsIris.Cell(4, 4).Value = _irisSummary.CanvasQty;
            wsIris.Cell(4, 5).Value = _irisSummary.DigitalUsbQty;
            wsIris.Cell(4, 6).Value = _irisSummary.Print10x10Qty;
            wsIris.Cell(4, 7).Value = _irisSummary.Print8x10Qty;
            wsIris.Cell(4, 8).Value = _irisSummary.Print10x13Qty;
            wsIris.Cell(4, 9).Value = _irisSummary.AddOnQty;
            wsIris.Cell(4, 10).FormulaA1 = "=SUM(C4:I4)";

            wsIris.Cell(5, 3).Value = Math.Round(_irisSummary.PremiumPkgAmt, 2);
            wsIris.Cell(5, 4).Value = Math.Round(_irisSummary.CanvasAmt, 2);
            wsIris.Cell(5, 5).Value = Math.Round(_irisSummary.DigitalUsbAmt, 2);
            wsIris.Cell(5, 6).Value = Math.Round(_irisSummary.Print10x10Amt, 2);
            wsIris.Cell(5, 7).Value = Math.Round(_irisSummary.Print8x10Amt, 2);
            wsIris.Cell(5, 8).Value = Math.Round(_irisSummary.Print10x13Amt, 2);
            wsIris.Cell(5, 9).Value = Math.Round(_irisSummary.AddOnAmt, 2);
            wsIris.Cell(5, 10).FormulaA1 = "=SUM(C5:I5)";

            wsIris.Range(3, 1, 5, 10).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            wsIris.Range(3, 1, 5, 10).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            wsIris.Columns().AdjustToContents();

            // TAB 5: ACCOUNTING_PIVOT
            var wsPivot = wb.Worksheets.Add("ACCOUNTING_PIVOT");
            wsPivot.Cell(1, 1).Value = voyage;
            wsPivot.Range(1, 1, 1, 5).Merge();
            wsPivot.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            wsPivot.Cell(1, 1).Style.Font.Bold = true;
            wsPivot.Cell(1, 1).Style.Font.FontSize = 14;

            wsPivot.Cell(2, 1).Value = voyage;
            wsPivot.Cell(2, 2).Value = "Sum of No.of Images";
            wsPivot.Cell(2, 3).Value = "Sum of Quantity Sold-Core";
            wsPivot.Cell(2, 4).Value = "Sum of Quantity Sold-Upsell";
            wsPivot.Cell(2, 5).Value = "Sum of Total Sale Amount";
            wsPivot.Range(2, 1, 2, 5).Style.Font.Bold = true;
            wsPivot.Range(2, 1, 2, 5).Style.Fill.BackgroundColor = xlMainColor;

            int pr = 3;
            foreach (var pg in _accountingRows.GroupBy(r => r.PhotographerName).OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
            {
                wsPivot.Cell(pr, 1).Value = pg.Key;
                wsPivot.Cell(pr, 2).Value = pg.Sum(x => x.Images);
                wsPivot.Cell(pr, 3).Value = pg.Sum(x => x.SoldCore);
                wsPivot.Cell(pr, 4).Value = pg.Sum(x => x.SoldUpsell);
                wsPivot.Cell(pr, 5).Value = Math.Round(pg.Sum(x => x.TotalAmount), 2);
                wsPivot.Range(pr, 1, pr, 5).Style.Font.Bold = true;
                wsPivot.Range(pr, 1, pr, 5).Style.Fill.BackgroundColor = xlSubColor;
                pr++;

                foreach (var cg in pg.GroupBy(x => x.CaptureType).OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
                {
                    wsPivot.Cell(pr, 1).Value = cg.Key;
                    wsPivot.Cell(pr, 2).Value = cg.Sum(x => x.Images);
                    wsPivot.Cell(pr, 3).Value = cg.Sum(x => x.SoldCore);
                    wsPivot.Cell(pr, 4).Value = cg.Sum(x => x.SoldUpsell);
                    wsPivot.Cell(pr, 5).Value = Math.Round(cg.Sum(x => x.TotalAmount), 2);
                    pr++;
                }
            }

            wsPivot.Cell(pr, 1).Value = "Grand Total";
            wsPivot.Cell(pr, 2).Value = _accountingRows.Sum(x => x.Images);
            wsPivot.Cell(pr, 3).Value = _accountingRows.Sum(x => x.SoldCore);
            wsPivot.Cell(pr, 4).Value = _accountingRows.Sum(x => x.SoldUpsell);
            wsPivot.Cell(pr, 5).Value = Math.Round(_accountingRows.Sum(x => x.TotalAmount), 2);
            wsPivot.Range(pr, 1, pr, 5).Style.Font.Bold = true;
            wsPivot.Range(pr, 1, pr, 5).Style.Fill.BackgroundColor = xlMainColor;
            wsPivot.Range(2, 1, pr, 5).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            wsPivot.Range(2, 1, pr, 5).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            wsPivot.Columns().AdjustToContents();

            // TAB 6: ORDER_PHOTO
            var wsPhoto = wb.Worksheets.Add("ORDER_PHOTO");
            for (int i = 0; i < _orderRowsForPhotoSales.Count; i++)
            {
                var r = _orderRowsForPhotoSales[i];
                wsPhoto.Cell(i + 1, 1).Value = r.FileName;
                wsPhoto.Cell(i + 1, 2).Value = SanitizeSpreadsheetText(r.PhotographerName);
            }
            wsPhoto.Columns().AdjustToContents();

            // TAB 7: ORDER_SALES
            var wsSales = wb.Worksheets.Add("ORDER_SALES");
            for (int i = 0; i < _orderRowsForPhotoSales.Count; i++)
            {
                var r = _orderRowsForPhotoSales[i];
                wsSales.Cell(i + 1, 1).Value = r.CabinNo;
                wsSales.Cell(i + 1, 2).Value = SanitizeSpreadsheetText(r.GuestName);
                wsSales.Cell(i + 1, 3).Value = r.FileName;
                wsSales.Cell(i + 1, 4).Value = r.MediaType;
            }
            wsSales.Columns().AdjustToContents();

            wb.SaveAs(filePath);
        }

        private void ExportOrderCsvs(string outDir)
        {
            string photoCsv = Path.Combine(outDir, "OrderReport-Photo.csv");
            string salesCsv = Path.Combine(outDir, "OrderReport-Sales.csv");

            var sbPhoto = new StringBuilder();
            var sbSales = new StringBuilder();

            foreach (var r in _orderRowsForPhotoSales)
            {
                sbPhoto.AppendLine($"{EscapeCsvField(r.FileName)},{EscapeCsvField(r.PhotographerName)}");
                sbSales.AppendLine($"{EscapeCsvField(r.CabinNo)},{EscapeCsvField(r.GuestName)},{EscapeCsvField(r.FileName)},{EscapeCsvField(r.MediaType)}");
            }

            File.WriteAllText(photoCsv, sbPhoto.ToString(), new UTF8Encoding(false));
            File.WriteAllText(salesCsv, sbSales.ToString(), new UTF8Encoding(false));
        }

        private List<PivotRenderLine> BuildPivotRenderLines(string voyage)
        {
            var lines = new List<PivotRenderLine>
            {
                new PivotRenderLine
                {
                    LineType = 1,
                    Label = voyage,
                    Col1 = "Sum of No.of Images",
                    Col2 = "Sum of Quantity Sold-Core",
                    Col3 = "Sum of Quantity Sold-Upsell",
                    Col4 = "Sum of Total Sale Amount"
                }
            };

            foreach (var pg in _accountingRows.GroupBy(r => r.PhotographerName).OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
            {
                lines.Add(new PivotRenderLine
                {
                    LineType = 2,
                    Label = pg.Key,
                    Col1 = pg.Sum(x => x.Images).ToString(),
                    Col2 = pg.Sum(x => x.SoldCore).ToString(),
                    Col3 = pg.Sum(x => x.SoldUpsell).ToString(),
                    Col4 = Math.Round(pg.Sum(x => x.TotalAmount), 2).ToString("0.##", CultureInfo.InvariantCulture)
                });

                foreach (var cg in pg.GroupBy(x => x.CaptureType).OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
                {
                    lines.Add(new PivotRenderLine
                    {
                        LineType = 3,
                        Label = cg.Key,
                        Col1 = cg.Sum(x => x.Images).ToString(),
                        Col2 = cg.Sum(x => x.SoldCore).ToString(),
                        Col3 = cg.Sum(x => x.SoldUpsell).ToString(),
                        Col4 = Math.Round(cg.Sum(x => x.TotalAmount), 2).ToString("0.##", CultureInfo.InvariantCulture)
                    });
                }
            }

            lines.Add(new PivotRenderLine
            {
                LineType = 4,
                Label = "Grand Total",
                Col1 = _accountingRows.Sum(x => x.Images).ToString(),
                Col2 = _accountingRows.Sum(x => x.SoldCore).ToString(),
                Col3 = _accountingRows.Sum(x => x.SoldUpsell).ToString(),
                Col4 = Math.Round(_accountingRows.Sum(x => x.TotalAmount), 2).ToString("0.##", CultureInfo.InvariantCulture)
            });

            return lines;
        }

        private int GetDynamicJpgHeightInches(int lineCount)
        {
            // Scales dynamically from 8x24 inches up to 8x36 inches at 300 DPI based on accounting report length
            if (lineCount <= 145) return 24;
            if (lineCount <= 185) return 30;
            if (lineCount <= 215) return 32;
            return 36;
        }

        private void GenerateSalesJpg(string jpgPath, string voyage, List<PivotRenderLine> prebuiltLines = null)
        {
            var lines = prebuiltLines ?? BuildPivotRenderLines(voyage);

            const int dpi = 300;
            int widthInches = 8;
            int heightInches = GetDynamicJpgHeightInches(lines.Count);

            int targetW = widthInches * dpi;   // 1600 px (8 inches @ 300 DPI)
            int targetH = heightInches * dpi;  // 4800 px to 7200 px (24 to 36 inches @ 300 DPI)

            using var bmp = new Bitmap(targetW, targetH, PixelFormat.Format24bppRgb);
            bmp.SetResolution(dpi, dpi);

            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                float marginX = 36f;
                float topTitleH = 92f;
                float bottomMargin = 28f;
                float availH = targetH - topTitleH - bottomMargin;
                float rowH = availH / Math.Max(1, lines.Count);

                using var titleFont = new Font("Calibri", 34f, FontStyle.Bold, GraphicsUnit.Pixel);
                var sfCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                g.DrawString(voyage, titleFont, Brushes.Black, new RectangleF(0, 10f, targetW, topTitleH - 14f), sfCenter);

                float tableW = targetW - (marginX * 2);
                float[] colW = { tableW * 0.42f, tableW * 0.145f, tableW * 0.145f, tableW * 0.145f, tableW * 0.145f };

                float fontSize = Math.Clamp(rowH * 0.64f, 12f, 28f);
                using var boldFont = new Font("Calibri", fontSize, FontStyle.Bold, GraphicsUnit.Pixel);
                using var regFont = new Font("Calibri", Math.Max(11f, fontSize * 0.96f), FontStyle.Regular, GraphicsUnit.Pixel);
                using var mainBrush = new SolidBrush(_activeMainColor);
                using var subBrush = new SolidBrush(_activeSubColor);
                using var gridPen = new Pen(Color.FromArgb(135, 135, 140), 1.5f);

                var sfLeft = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center, FormatFlags = StringFormatFlags.NoWrap };
                var sfRight = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center, FormatFlags = StringFormatFlags.NoWrap };

                float y = topTitleH;
                foreach (var ln in lines)
                {
                    var rowRect = new RectangleF(marginX, y, tableW, rowH);
                    if (ln.LineType == 1 || ln.LineType == 4)
                        g.FillRectangle(mainBrush, rowRect);
                    else if (ln.LineType == 2)
                        g.FillRectangle(subBrush, rowRect);

                    Font f = (ln.LineType == 3) ? regFont : boldFont;
                    float cx = marginX;
                    string[] vals = { ln.Label, ln.Col1, ln.Col2, ln.Col3, ln.Col4 };

                    for (int c = 0; c < 5; c++)
                    {
                        g.DrawRectangle(gridPen, cx, y, colW[c], rowH);
                        var padRect = new RectangleF(cx + (c == 0 && ln.LineType == 3 ? 18f : 8f), y, colW[c] - 16f, rowH);
                        g.DrawString(vals[c], f, Brushes.Black, padRect, c == 0 ? sfLeft : sfRight);
                        cx += colW[c];
                    }

                    y += rowH;
                }
            }

            var jpgEncoder = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);
            var encParams = new EncoderParameters(1);
            encParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 96L);
            bmp.Save(jpgPath, jpgEncoder, encParams);
        }

        private void GenerateAccountingPdf(string pdfPath, string voyage, List<PivotRenderLine> prebuiltLines = null)
        {
            var lines = prebuiltLines ?? BuildPivotRenderLines(voyage);
            int lineIndex = 0;

            using var doc = new PrintDocument();
            doc.PrinterSettings.PrinterName = "Microsoft Print to PDF";
            if (!doc.PrinterSettings.IsValid)
            {
                return;
            }
            doc.PrinterSettings.PrintToFile = true;
            doc.PrinterSettings.PrintFileName = pdfPath;
            doc.DefaultPageSettings.PaperSize = new PaperSize("Letter", 850, 1100);
            doc.DefaultPageSettings.Margins = new Margins(36, 36, 36, 36);

            doc.PrintPage += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                float x0 = e.MarginBounds.Left;
                float y = e.MarginBounds.Top;
                float totalW = e.MarginBounds.Width;
                float bottomLimit = e.MarginBounds.Bottom;

                if (lineIndex == 0)
                {
                    using var titleFont = new Font("Calibri", 14f, FontStyle.Bold);
                    var sfCenter = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                    g.DrawString(voyage, titleFont, Brushes.Black, new RectangleF(x0, y, totalW, 26f), sfCenter);
                    y += 30f;
                }

                float[] colW = { totalW * 0.40f, totalW * 0.15f, totalW * 0.15f, totalW * 0.15f, totalW * 0.15f };
                using var boldFont = new Font("Calibri", 9f, FontStyle.Bold);
                using var regFont = new Font("Calibri", 9f, FontStyle.Regular);
                using var mainBrush = new SolidBrush(_activeMainColor);
                using var subBrush = new SolidBrush(_activeSubColor);
                using var gridPen = new Pen(Color.FromArgb(120, 120, 120), 0.75f);

                var sfLeft = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };
                var sfRight = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center };

                while (lineIndex < lines.Count)
                {
                    var ln = lines[lineIndex];
                    float rowH = (ln.LineType == 1) ? 28f : 15.2f;

                    if (y + rowH > bottomLimit)
                    {
                        e.HasMorePages = true;
                        return;
                    }

                    var rowRect = new RectangleF(x0, y, totalW, rowH);
                    if (ln.LineType == 1 || ln.LineType == 4)
                        g.FillRectangle(mainBrush, rowRect);
                    else if (ln.LineType == 2)
                        g.FillRectangle(subBrush, rowRect);

                    Font f = (ln.LineType == 3) ? regFont : boldFont;
                    float cx = x0;
                    string[] vals = { ln.Label, ln.Col1, ln.Col2, ln.Col3, ln.Col4 };

                    for (int c = 0; c < 5; c++)
                    {
                        g.DrawRectangle(gridPen, cx, y, colW[c], rowH);
                        var padRect = new RectangleF(cx + (c == 0 && ln.LineType == 3 ? 12f : 4f), y, colW[c] - 8f, rowH);
                        g.DrawString(vals[c], f, Brushes.Black, padRect, c == 0 ? sfLeft : sfRight);
                        cx += colW[c];
                    }

                    y += rowH;
                    lineIndex++;
                }

                e.HasMorePages = false;
            };

            doc.Print();
        }
    }
}
