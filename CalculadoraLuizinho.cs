using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace CalculadoraLuizinho
{
    public class CalculationRecord
    {
        public int Index { get; set; }
        public DateTime Timestamp { get; set; }
        public string Expression { get; set; }
        public string ResultFormatted { get; set; }
        public double ResultValue { get; set; }

        public override string ToString()
        {
            return string.Format("[{0:HH:mm:ss}]  {1} {2}", Timestamp, Expression, ResultFormatted);
        }
    }

    public class UnitDef
    {
        public string Name { get; set; }
        public string Symbol { get; set; }
        public double FactorToBase { get; set; } // For linear conversions

        public UnitDef(string name, string symbol, double factorToBase)
        {
            Name = name;
            Symbol = symbol;
            FactorToBase = factorToBase;
        }

        public override string ToString()
        {
            return string.Format("{0} ({1})", Name, Symbol);
        }
    }

    public class MainForm : Form
    {
        // DWM Dark Title Bar API (Windows 10/11)
        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        // Theme Colors (Windows 10/11 Dark Calculator)
        private readonly Color BgDark = Color.FromArgb(32, 32, 32);          // #202020
        private readonly Color DrawerBg = Color.FromArgb(43, 43, 43);        // #2B2B2B
        private readonly Color DrawerHover = Color.FromArgb(60, 60, 60);     // #3C3C3C
        private readonly Color BtnFuncBg = Color.FromArgb(50, 50, 50);       // #323232
        private readonly Color BtnNumBg = Color.FromArgb(59, 59, 59);        // #3B3B3B
        private readonly Color BtnHoverBg = Color.FromArgb(72, 72, 72);      // #484848
        private readonly Color BtnEqualBg = Color.FromArgb(118, 185, 237);   // #76B9ED
        private readonly Color BtnEqualHover = Color.FromArgb(100, 168, 220);
        private readonly Color TextWhite = Color.FromArgb(255, 255, 255);
        private readonly Color TextMuted = Color.FromArgb(166, 166, 166);
        private readonly Color TextDisabled = Color.FromArgb(95, 95, 95);
        private readonly Color AccentBlue = Color.FromArgb(118, 185, 237);
        private readonly Color DangerRed = Color.FromArgb(215, 70, 70);

        // Header & Main UI Controls
        private Button btnMenu;
        private Label lblTitle;
        private Label lblModeSubtitle;
        private Button btnAlwaysOnTop;
        private Button btnToggleHistory;
        private Label lblExpression;
        private Label lblDisplay;
        private Button[] memButtons = new Button[6];
        private Button btnMc, btnMr, btnMPlus, btnMMinus, btnMs, btnMv;
        private Button btnPrintReport, btnClearHistory;
        private Panel pnlKeypad;
        private List<Button> keypadButtons = new List<Button>();

        // Special Mode Panels (Converter / Date / Programmer info)
        private Panel pnlModeExtra;
        private ComboBox cmbUnitFrom;
        private ComboBox cmbUnitTo;
        private Label lblConvertedPreview;
        private DateTimePicker dtpStart;
        private DateTimePicker dtpEnd;
        private Label lblDateDiffResult;

        // Side History Drawer Panel
        private Panel pnlHistorySide;
        private Label lblHistHeader;
        private Label lblHistoryCount;
        private ListBox lstHistory;
        private Button btnSidePrint;
        private Button btnSideClear;

        // Navigation Menu Drawer ("Menu Opções" igual da original)
        private Panel pnlMenuOverlay;
        private Panel pnlNavDrawer;
        private Panel pnlDrawerScroll;
        private Button btnDrawerHamburger;
        private Label lblDrawerBrand;
        private List<Control> navMenuItems = new List<Control>();
        private bool menuOpen = false;

        // Mode & Calculator State
        private string currentMode = "Padrão";
        private bool isAlwaysOnTop = false;
        private Size normalSizeBeforeCompact = new Size(360, 580);
        private bool degMode = true; // For Scientific mode (DEG vs RAD)

        private string currentInput = "0";
        private double accumulator = 0;
        private string pendingOp = "";
        private bool startNewInput = true;
        private bool lastWasEquals = false;
        private double lastOperand = 0;
        private string lastBinaryOp = "";

        // Memory & History State
        private List<double> memoryStack = new List<double>();
        private List<CalculationRecord> calculationHistory = new List<CalculationRecord>();
        private bool historyUserToggled = false;
        private bool historyAutoDocked = false;
        private readonly CultureInfo ptBr = new CultureInfo("pt-BR");
        private readonly string historyFilePath;
        private ToolTip mainToolTip;

        // Converter Definitions
        private Dictionary<string, List<UnitDef>> converterUnits = new Dictionary<string, List<UnitDef>>();

        // Printing state
        private int printCurrentIndex = 0;

        public MainForm()
        {
            historyFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "historico_calculos.txt");
            InitConverterDatabase();
            InitializeComponent();
            EnableDarkTitleBar();
            LoadSavedHistory();
            UpdateMemoryButtonsState();
            UpdateHistoryBadge();
            BuildKeypadForCurrentMode();
            PerformResponsiveLayout();
        }

        private void EnableDarkTitleBar()
        {
            try
            {
                int useDark = 1;
                if (DwmSetWindowAttribute(this.Handle, 20, ref useDark, sizeof(int)) != 0)
                {
                    DwmSetWindowAttribute(this.Handle, 19, ref useDark, sizeof(int));
                }
            }
            catch { }
        }

        private void InitConverterDatabase()
        {
            converterUnits["Moeda"] = new List<UnitDef>
            {
                new UnitDef("Real Brasileiro", "BRL", 1.0),
                new UnitDef("Dólar Americano", "USD", 5.70),
                new UnitDef("Euro", "EUR", 6.15),
                new UnitDef("Libra Esterlina", "GBP", 7.35),
                new UnitDef("Iene Japonês", "JPY", 0.038),
                new UnitDef("Peso Argentino", "ARS", 0.0058)
            };
            converterUnits["Volume"] = new List<UnitDef>
            {
                new UnitDef("Mililitros", "mL", 0.001),
                new UnitDef("Litros", "L", 1.0),
                new UnitDef("Metros cúbicos", "m³", 1000.0),
                new UnitDef("Galões (EUA)", "gal", 3.785411784),
                new UnitDef("Onças líquidas (EUA)", "fl oz", 0.0295735)
            };
            converterUnits["Comprimento"] = new List<UnitDef>
            {
                new UnitDef("Milímetros", "mm", 0.001),
                new UnitDef("Centímetros", "cm", 0.01),
                new UnitDef("Metros", "m", 1.0),
                new UnitDef("Quilômetros", "km", 1000.0),
                new UnitDef("Polegadas", "in", 0.0254),
                new UnitDef("Pés", "ft", 0.3048),
                new UnitDef("Jardas", "yd", 0.9144),
                new UnitDef("Milhas", "mi", 1609.344)
            };
            converterUnits["Peso e Massa"] = new List<UnitDef>
            {
                new UnitDef("Miligramas", "mg", 0.000001),
                new UnitDef("Gramas", "g", 0.001),
                new UnitDef("Quilogramas", "kg", 1.0),
                new UnitDef("Toneladas Métricas", "t", 1000.0),
                new UnitDef("Onças", "oz", 0.0283495),
                new UnitDef("Libras", "lb", 0.45359237)
            };
            converterUnits["Temperatura"] = new List<UnitDef>
            {
                new UnitDef("Celsius", "°C", 1.0),
                new UnitDef("Fahrenheit", "°F", 1.0),
                new UnitDef("Kelvin", "K", 1.0)
            };
            converterUnits["Energia"] = new List<UnitDef>
            {
                new UnitDef("Joules", "J", 1.0),
                new UnitDef("Quilojoules", "kJ", 1000.0),
                new UnitDef("Calorias alimentares", "kcal", 4184.0),
                new UnitDef("Quilowatt-hora", "kWh", 3600000.0),
                new UnitDef("BTU", "BTU", 1055.056)
            };
            converterUnits["Área"] = new List<UnitDef>
            {
                new UnitDef("Centímetros quadrados", "cm²", 0.0001),
                new UnitDef("Metros quadrados", "m²", 1.0),
                new UnitDef("Hectares", "ha", 10000.0),
                new UnitDef("Quilômetros quadrados", "km²", 1000000.0),
                new UnitDef("Pés quadrados", "ft²", 0.092903),
                new UnitDef("Acres", "ac", 4046.856)
            };
            converterUnits["Velocidade"] = new List<UnitDef>
            {
                new UnitDef("Metros por segundo", "m/s", 1.0),
                new UnitDef("Quilômetros por hora", "km/h", 0.27777778),
                new UnitDef("Milhas por hora", "mph", 0.44704),
                new UnitDef("Nós", "kn", 0.514444),
                new UnitDef("Mach", "Ma", 340.3)
            };
            converterUnits["Tempo"] = new List<UnitDef>
            {
                new UnitDef("Milissegundos", "ms", 0.001),
                new UnitDef("Segundos", "s", 1.0),
                new UnitDef("Minutos", "min", 60.0),
                new UnitDef("Horas", "h", 3600.0),
                new UnitDef("Dias", "d", 86400.0),
                new UnitDef("Semanas", "sem", 604800.0),
                new UnitDef("Anos (365d)", "anos", 31536000.0)
            };
            converterUnits["Potência"] = new List<UnitDef>
            {
                new UnitDef("Watts", "W", 1.0),
                new UnitDef("Quilowatts", "kW", 1000.0),
                new UnitDef("Cavalos-vapor (cv/hp)", "hp", 745.7),
                new UnitDef("BTUs/minuto", "BTU/min", 17.58426)
            };
            converterUnits["Dados"] = new List<UnitDef>
            {
                new UnitDef("Bits", "bit", 0.125),
                new UnitDef("Bytes", "B", 1.0),
                new UnitDef("Kilobytes", "KB", 1024.0),
                new UnitDef("Megabytes", "MB", 1048576.0),
                new UnitDef("Gigabytes", "GB", 1073741824.0),
                new UnitDef("Terabytes", "TB", 1099511627776.0)
            };
            converterUnits["Pressão"] = new List<UnitDef>
            {
                new UnitDef("Pascals", "Pa", 1.0),
                new UnitDef("Quilopascals", "kPa", 1000.0),
                new UnitDef("Bares", "bar", 100000.0),
                new UnitDef("Atmosferas", "atm", 101325.0),
                new UnitDef("Libras por pol²", "PSI", 6894.757),
                new UnitDef("Milímetros de mercúrio", "mmHg", 133.322)
            };
            converterUnits["Ângulo"] = new List<UnitDef>
            {
                new UnitDef("Graus", "°", 1.0),
                new UnitDef("Radianos", "rad", 57.295779513),
                new UnitDef("Gradianos", "grad", 0.9)
            };
        }

        private void InitializeComponent()
        {
            this.Text = "Calculadora — Created by LuizinhoSouzaK";
            this.BackColor = BgDark;
            this.ForeColor = TextWhite;
            this.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            this.ClientSize = new Size(360, 580);
            this.MinimumSize = new Size(280, 440);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.KeyPreview = true;
            this.DoubleBuffered = true;

            try
            {
                this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch { }

            mainToolTip = new ToolTip();

            // 1. Header Bar ("☰  Created by LuizinhoSouzaK  [⧉ Sobrepor]  [🕒]")
            btnMenu = CreateFlatIconButton("☰", new Point(8, 8), new Size(38, 36), 13F);
            mainToolTip.SetToolTip(btnMenu, "Abrir Navegação (Menu Opções)");
            btnMenu.Click += (s, e) => ToggleNavMenu();

            lblTitle = new Label
            {
                Text = "Created by LuizinhoSouzaK",
                Font = new Font("Segoe UI Semibold", 11F, FontStyle.Bold),
                ForeColor = TextWhite,
                Location = new Point(48, 6),
                Size = new Size(210, 22),
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };

            lblModeSubtitle = new Label
            {
                Text = "Padrão",
                Font = new Font("Segoe UI", 8.5F, FontStyle.Regular),
                ForeColor = AccentBlue,
                Location = new Point(48, 26),
                Size = new Size(210, 16),
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };

            // Botão "Manter no topo" / Sobrepor outros apps (igual ao da Calculadora original do Windows!)
            btnAlwaysOnTop = new Button
            {
                Text = "",
                BackColor = BgDark,
                ForeColor = TextWhite,
                FlatStyle = FlatStyle.Flat,
                Location = new Point(260, 8),
                Size = new Size(38, 36),
                TabStop = false,
                Cursor = Cursors.Hand
            };
            btnAlwaysOnTop.FlatAppearance.BorderSize = 0;
            btnAlwaysOnTop.FlatAppearance.MouseOverBackColor = BtnFuncBg;
            btnAlwaysOnTop.Paint += BtnAlwaysOnTop_Paint;
            btnAlwaysOnTop.Click += (s, e) => ToggleAlwaysOnTop();
            mainToolTip.SetToolTip(btnAlwaysOnTop, "Manter no topo (Sobrepor outros apps) — Alt+Seta Acima");

            btnToggleHistory = CreateFlatIconButton("🕒", new Point(302, 8), new Size(38, 36), 12F);
            mainToolTip.SetToolTip(btnToggleHistory, "Mostrar/Ocultar Painel de Histórico de Cálculos (H)");
            btnToggleHistory.Click += (s, e) => ToggleHistoryPanel();

            this.Controls.Add(btnMenu);
            this.Controls.Add(lblTitle);
            this.Controls.Add(lblModeSubtitle);
            this.Controls.Add(btnAlwaysOnTop);
            this.Controls.Add(btnToggleHistory);

            // 2. Expression & Main Display
            lblExpression = new Label
            {
                Text = "",
                Font = new Font("Segoe UI", 10.5F, FontStyle.Regular),
                ForeColor = TextMuted,
                Location = new Point(10, 48),
                Size = new Size(340, 24),
                TextAlign = ContentAlignment.MiddleRight,
                AutoEllipsis = true
            };

            lblDisplay = new Label
            {
                Text = "0",
                Font = new Font("Segoe UI Semibold", 34F, FontStyle.Bold),
                ForeColor = TextWhite,
                Location = new Point(6, 72),
                Size = new Size(348, 66),
                TextAlign = ContentAlignment.MiddleRight
            };

            this.Controls.Add(lblExpression);
            this.Controls.Add(lblDisplay);

            // 2b. Extra Panel for Unit Converter or Date Calculator
            pnlModeExtra = new Panel
            {
                BackColor = Color.FromArgb(38, 38, 38),
                Visible = false
            };

            cmbUnitFrom = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BtnFuncBg,
                ForeColor = TextWhite,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F)
            };
            cmbUnitFrom.SelectedIndexChanged += (s, e) => UpdateConversionPreview();

            cmbUnitTo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = BtnFuncBg,
                ForeColor = TextWhite,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9F)
            };
            cmbUnitTo.SelectedIndexChanged += (s, e) => UpdateConversionPreview();

            lblConvertedPreview = new Label
            {
                Text = "= 0",
                Font = new Font("Segoe UI Semibold", 11F, FontStyle.Bold),
                ForeColor = AccentBlue,
                TextAlign = ContentAlignment.MiddleRight
            };

            dtpStart = new DateTimePicker
            {
                Format = DateTimePickerFormat.Short,
                Font = new Font("Segoe UI", 9.5F),
                Visible = false
            };
            dtpStart.ValueChanged += (s, e) => UpdateDateCalculation();

            dtpEnd = new DateTimePicker
            {
                Format = DateTimePickerFormat.Short,
                Font = new Font("Segoe UI", 9.5F),
                Visible = false
            };
            dtpEnd.ValueChanged += (s, e) => UpdateDateCalculation();

            lblDateDiffResult = new Label
            {
                Text = "Mesmas datas (0 dias)",
                Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
                ForeColor = AccentBlue,
                TextAlign = ContentAlignment.MiddleCenter,
                Visible = false
            };

            pnlModeExtra.Controls.Add(cmbUnitFrom);
            pnlModeExtra.Controls.Add(cmbUnitTo);
            pnlModeExtra.Controls.Add(lblConvertedPreview);
            pnlModeExtra.Controls.Add(dtpStart);
            pnlModeExtra.Controls.Add(dtpEnd);
            pnlModeExtra.Controls.Add(lblDateDiffResult);
            this.Controls.Add(pnlModeExtra);

            // 3. Memory Row (MC, MR, M+, M-, MS, Mv)
            string[] memLabels = { "MC", "MR", "M+", "M−", "MS", "M▾" };
            for (int i = 0; i < 6; i++)
            {
                Button b = new Button
                {
                    Text = memLabels[i],
                    Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
                    ForeColor = TextWhite,
                    BackColor = BgDark,
                    FlatStyle = FlatStyle.Flat,
                    TabStop = false,
                    Cursor = Cursors.Hand
                };
                b.FlatAppearance.BorderSize = 0;
                b.FlatAppearance.MouseOverBackColor = BtnFuncBg;
                memButtons[i] = b;
                this.Controls.Add(b);
            }
            btnMc = memButtons[0];
            btnMr = memButtons[1];
            btnMPlus = memButtons[2];
            btnMMinus = memButtons[3];
            btnMs = memButtons[4];
            btnMv = memButtons[5];

            btnMc.Click += (s, e) => { memoryStack.Clear(); UpdateMemoryButtonsState(); };
            btnMr.Click += (s, e) => { if (memoryStack.Count > 0) { SetDisplayValue(memoryStack[0]); startNewInput = true; } };
            btnMPlus.Click += (s, e) => { ModifyMemory(GetDisplayValue(), true); };
            btnMMinus.Click += (s, e) => { ModifyMemory(GetDisplayValue(), false); };
            btnMs.Click += (s, e) => { memoryStack.Insert(0, GetDisplayValue()); UpdateMemoryButtonsState(); startNewInput = true; };
            btnMv.Click += (s, e) => ShowMemoryPopup();

            // 4. History Report & Clear History Bar
            btnPrintReport = new Button
            {
                Text = "🖨 Imprimir Relatório (0)",
                Font = new Font("Segoe UI Semibold", 8.8F, FontStyle.Bold),
                ForeColor = TextWhite,
                BackColor = Color.FromArgb(42, 68, 92),
                FlatStyle = FlatStyle.Flat,
                TabStop = false,
                Cursor = Cursors.Hand
            };
            btnPrintReport.FlatAppearance.BorderColor = AccentBlue;
            btnPrintReport.FlatAppearance.BorderSize = 1;
            btnPrintReport.FlatAppearance.MouseOverBackColor = Color.FromArgb(56, 90, 122);
            btnPrintReport.Click += (s, e) => ShowPrintReportOptions();

            btnClearHistory = new Button
            {
                Text = "🗑 Limpar Histórico",
                Font = new Font("Segoe UI Semibold", 8.8F, FontStyle.Bold),
                ForeColor = TextWhite,
                BackColor = Color.FromArgb(75, 38, 38),
                FlatStyle = FlatStyle.Flat,
                TabStop = false,
                Cursor = Cursors.Hand
            };
            btnClearHistory.FlatAppearance.BorderColor = DangerRed;
            btnClearHistory.FlatAppearance.BorderSize = 1;
            btnClearHistory.FlatAppearance.MouseOverBackColor = Color.FromArgb(105, 48, 48);
            btnClearHistory.Click += (s, e) => ClearHistoryAction();

            this.Controls.Add(btnPrintReport);
            this.Controls.Add(btnClearHistory);

            // 5. Responsive Keypad Container Panel
            pnlKeypad = new Panel
            {
                BackColor = BgDark
            };
            this.Controls.Add(pnlKeypad);

            // 6. Side History Drawer Panel
            pnlHistorySide = new Panel
            {
                BackColor = Color.FromArgb(26, 26, 26),
                Visible = false
            };

            lblHistHeader = new Label
            {
                Text = "📜 Histórico na Memória",
                Font = new Font("Segoe UI Semibold", 11F, FontStyle.Bold),
                ForeColor = AccentBlue
            };

            lblHistoryCount = new Label
            {
                Text = "Nenhum cálculo registrado ainda.",
                Font = new Font("Segoe UI", 8.5F, FontStyle.Regular),
                ForeColor = TextMuted
            };

            lstHistory = new ListBox
            {
                BackColor = Color.FromArgb(38, 38, 38),
                ForeColor = TextWhite,
                Font = new Font("Consolas", 9.5F, FontStyle.Regular),
                BorderStyle = BorderStyle.FixedSingle,
                HorizontalScrollbar = true,
                TabStop = false
            };
            lstHistory.DoubleClick += (s, e) =>
            {
                if (lstHistory.SelectedIndex >= 0 && lstHistory.SelectedIndex < calculationHistory.Count)
                {
                    CalculationRecord rec = calculationHistory[lstHistory.SelectedIndex];
                    SetDisplayValue(rec.ResultValue);
                    startNewInput = true;
                }
            };

            btnSidePrint = new Button
            {
                Text = "🖨 Imprimir",
                Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
                ForeColor = TextWhite,
                BackColor = Color.FromArgb(42, 68, 92),
                FlatStyle = FlatStyle.Flat,
                TabStop = false,
                Cursor = Cursors.Hand
            };
            btnSidePrint.FlatAppearance.BorderColor = AccentBlue;
            btnSidePrint.Click += (s, e) => ShowPrintReportOptions();

            btnSideClear = new Button
            {
                Text = "🗑 Limpar Tudo",
                Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
                ForeColor = TextWhite,
                BackColor = Color.FromArgb(75, 38, 38),
                FlatStyle = FlatStyle.Flat,
                TabStop = false,
                Cursor = Cursors.Hand
            };
            btnSideClear.FlatAppearance.BorderColor = DangerRed;
            btnSideClear.Click += (s, e) => ClearHistoryAction();

            pnlHistorySide.Controls.Add(lblHistHeader);
            pnlHistorySide.Controls.Add(lblHistoryCount);
            pnlHistorySide.Controls.Add(lstHistory);
            pnlHistorySide.Controls.Add(btnSidePrint);
            pnlHistorySide.Controls.Add(btnSideClear);
            this.Controls.Add(pnlHistorySide);

            // 7. Navigation Drawer ("Menu Opções" igual ao da Calculadora Original do Windows)
            BuildNavigationDrawer();

            this.Shown += (s, e) => { this.ActiveControl = null; };
            this.KeyDown += MainForm_KeyDown;
            this.KeyPress += MainForm_KeyPress;
            this.Resize += (s, e) => PerformResponsiveLayout();
        }

        // Custom-paint the Windows 10/11 "Keep on Top" (Sobrepor outros apps) icon
        private void BtnAlwaysOnTop_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            Color iconColor = isAlwaysOnTop ? AccentBlue : TextWhite;
            int w = btnAlwaysOnTop.Width;
            int h = btnAlwaysOnTop.Height;

            if (isAlwaysOnTop)
            {
                using (Pen borderPen = new Pen(AccentBlue, 1.5f))
                {
                    g.DrawRectangle(borderPen, 2, 2, w - 5, h - 5);
                }
            }

            int boxX = (w - 18) / 2;
            int boxY = (h - 16) / 2;

            using (Pen p = new Pen(iconColor, 1.5f))
            {
                if (!isAlwaysOnTop)
                {
                    // Outer window frame (dashed/muted bottom-right + solid top-left overlay window + arrow pointing to top-left)
                    g.DrawRectangle(p, boxX, boxY, 17, 15);
                    using (SolidBrush fill = new SolidBrush(iconColor))
                    {
                        g.FillRectangle(fill, boxX, boxY, 9, 8);
                    }
                    // Diagonal arrow from bottom-right pointing toward top-left mini window
                    g.DrawLine(p, boxX + 15, boxY + 13, boxX + 10, boxY + 8);
                    g.DrawLine(p, boxX + 10, boxY + 8, boxX + 14, boxY + 8);
                    g.DrawLine(p, boxX + 10, boxY + 8, boxX + 10, boxY + 12);
                }
                else
                {
                    // Active "Restore / Always on top active" icon: full frame with arrow pointing bottom-right
                    g.DrawRectangle(p, boxX, boxY, 17, 15);
                    using (SolidBrush fill = new SolidBrush(AccentBlue))
                    {
                        g.FillRectangle(fill, boxX + 1, boxY + 1, 8, 7);
                    }
                    g.DrawLine(p, boxX + 9, boxY + 8, boxX + 14, boxY + 13);
                    g.DrawLine(p, boxX + 14, boxY + 13, boxX + 10, boxY + 13);
                    g.DrawLine(p, boxX + 14, boxY + 13, boxX + 14, boxY + 9);
                }
            }
        }

        private void ToggleAlwaysOnTop()
        {
            isAlwaysOnTop = !isAlwaysOnTop;
            this.TopMost = isAlwaysOnTop;

            if (isAlwaysOnTop)
            {
                if (this.WindowState == FormWindowState.Normal)
                {
                    normalSizeBeforeCompact = this.ClientSize;
                    // Compact overlay size while still remaining resizable
                    this.ClientSize = new Size(Math.Max(310, Math.Min(340, this.ClientSize.Width)), Math.Max(480, Math.Min(520, this.ClientSize.Height)));
                }
                mainToolTip.SetToolTip(btnAlwaysOnTop, "Sair do modo Sempre no Topo (Sobrepor apps ATIVADO) — Alt+Seta Abaixo");
                lblModeSubtitle.Text = currentMode + " • 📌 Sobrepondo Apps";
            }
            else
            {
                if (this.WindowState == FormWindowState.Normal && normalSizeBeforeCompact.Width > 0)
                {
                    this.ClientSize = normalSizeBeforeCompact;
                }
                mainToolTip.SetToolTip(btnAlwaysOnTop, "Manter no topo (Sobrepor outros apps) — Alt+Seta Acima");
                lblModeSubtitle.Text = currentMode;
            }

            btnAlwaysOnTop.Invalidate();
            PerformResponsiveLayout();
        }

        private void BuildNavigationDrawer()
        {
            // Click-away overlay behind the drawer
            pnlMenuOverlay = new Panel
            {
                BackColor = Color.FromArgb(20, 20, 20),
                Visible = false
            };
            pnlMenuOverlay.Click += (s, e) => CloseNavMenu();

            // Left slide-out drawer matching original Windows Calculator
            pnlNavDrawer = new Panel
            {
                BackColor = DrawerBg,
                Visible = false,
                Width = 264
            };

            // Drawer Top Bar
            btnDrawerHamburger = CreateFlatIconButton("☰", new Point(8, 8), new Size(38, 36), 13F);
            btnDrawerHamburger.BackColor = DrawerBg;
            btnDrawerHamburger.Click += (s, e) => CloseNavMenu();

            lblDrawerBrand = new Label
            {
                Text = "Created by LuizinhoSouzaK",
                Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold),
                ForeColor = AccentBlue,
                Location = new Point(50, 10),
                Size = new Size(204, 32),
                TextAlign = ContentAlignment.MiddleLeft
            };

            // Scrollable middle list of Original Calculator Options
            pnlDrawerScroll = new Panel
            {
                Location = new Point(0, 48),
                Size = new Size(264, 450),
                AutoScroll = true,
                BackColor = DrawerBg
            };

            int y = 4;
            y = AddDrawerSectionHeader("Calculadora", y);
            y = AddDrawerOption("🔢", "Padrão", y);
            y = AddDrawerOption("🧪", "Científica", y);
            y = AddDrawerOption("📈", "Representação Gráfica", y);
            y = AddDrawerOption("💻", "Programador", y);
            y = AddDrawerOption("📅", "Cálculo de Data", y);

            y += 8;
            y = AddDrawerSectionHeader("Conversor", y);
            y = AddDrawerOption("💱", "Moeda", y);
            y = AddDrawerOption("🧊", "Volume", y);
            y = AddDrawerOption("📏", "Comprimento", y);
            y = AddDrawerOption("⚖️", "Peso e Massa", y);
            y = AddDrawerOption("🌡️", "Temperatura", y);
            y = AddDrawerOption("⚡", "Energia", y);
            y = AddDrawerOption("📐", "Área", y);
            y = AddDrawerOption("🚀", "Velocidade", y);
            y = AddDrawerOption("⏱️", "Tempo", y);
            y = AddDrawerOption("🔌", "Potência", y);
            y = AddDrawerOption("💾", "Dados", y);
            y = AddDrawerOption("🧭", "Pressão", y);
            y = AddDrawerOption("📐", "Ângulo", y);

            // Bottom Footer inside Drawer ("Sobre / Configurações")
            Panel pnlDrawerFooter = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 76,
                BackColor = Color.FromArgb(36, 36, 36)
            };

            Button btnDrawerPrint = new Button
            {
                Text = "  🖨️   Imprimir Relatório de Cálculos",
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                ForeColor = TextWhite,
                BackColor = Color.FromArgb(36, 36, 36),
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleLeft,
                Location = new Point(0, 2),
                Size = new Size(264, 36),
                Cursor = Cursors.Hand,
                TabStop = false
            };
            btnDrawerPrint.FlatAppearance.BorderSize = 0;
            btnDrawerPrint.FlatAppearance.MouseOverBackColor = DrawerHover;
            btnDrawerPrint.Click += (s, e) =>
            {
                CloseNavMenu();
                ShowPrintReportOptions();
            };

            Button btnDrawerAbout = new Button
            {
                Text = "  ⚙️   Sobre — Created by LuizinhoSouzaK",
                Font = new Font("Segoe UI Semibold", 9.2F, FontStyle.Bold),
                ForeColor = AccentBlue,
                BackColor = Color.FromArgb(36, 36, 36),
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleLeft,
                Location = new Point(0, 38),
                Size = new Size(264, 36),
                Cursor = Cursors.Hand,
                TabStop = false
            };
            btnDrawerAbout.FlatAppearance.BorderSize = 0;
            btnDrawerAbout.FlatAppearance.MouseOverBackColor = DrawerHover;
            btnDrawerAbout.Click += (s, e) =>
            {
                CloseNavMenu();
                ShowAboutDialog();
            };

            pnlDrawerFooter.Controls.Add(btnDrawerPrint);
            pnlDrawerFooter.Controls.Add(btnDrawerAbout);

            pnlNavDrawer.Controls.Add(btnDrawerHamburger);
            pnlNavDrawer.Controls.Add(lblDrawerBrand);
            pnlNavDrawer.Controls.Add(pnlDrawerScroll);
            pnlNavDrawer.Controls.Add(pnlDrawerFooter);

            this.Controls.Add(pnlNavDrawer);
            this.Controls.Add(pnlMenuOverlay);
        }

        private int AddDrawerSectionHeader(string title, int y)
        {
            Label lbl = new Label
            {
                Text = title,
                Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
                ForeColor = TextMuted,
                Location = new Point(16, y),
                Size = new Size(220, 26),
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnlDrawerScroll.Controls.Add(lbl);
            return y + 26;
        }

        private int AddDrawerOption(string icon, string modeName, int y)
        {
            Button btn = new Button
            {
                Tag = modeName,
                Text = string.Format("  {0}   {1}", icon, modeName),
                Font = new Font("Segoe UI", 9.8F, modeName == currentMode ? FontStyle.Bold : FontStyle.Regular),
                ForeColor = modeName == currentMode ? AccentBlue : TextWhite,
                BackColor = modeName == currentMode ? DrawerHover : DrawerBg,
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleLeft,
                Location = new Point(4, y),
                Size = new Size(236, 36),
                Cursor = Cursors.Hand,
                TabStop = false
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = DrawerHover;
            btn.Paint += (s, e) =>
            {
                if ((string)btn.Tag == currentMode)
                {
                    using (SolidBrush b = new SolidBrush(AccentBlue))
                    {
                        e.Graphics.FillRectangle(b, 0, 8, 4, btn.Height - 16);
                    }
                }
            };
            btn.Click += (s, e) => SelectCalculatorMode(modeName);
            pnlDrawerScroll.Controls.Add(btn);
            navMenuItems.Add(btn);
            return y + 38;
        }

        private void ToggleNavMenu()
        {
            if (menuOpen)
                CloseNavMenu();
            else
                OpenNavMenu();
        }

        private void OpenNavMenu()
        {
            menuOpen = true;
            foreach (Control c in navMenuItems)
            {
                Button b = c as Button;
                if (b != null)
                {
                    bool active = ((string)b.Tag == currentMode);
                    b.Font = new Font("Segoe UI", 9.8F, active ? FontStyle.Bold : FontStyle.Regular);
                    b.ForeColor = active ? AccentBlue : TextWhite;
                    b.BackColor = active ? DrawerHover : DrawerBg;
                    b.Invalidate();
                }
            }

            pnlNavDrawer.Location = new Point(0, 0);
            pnlNavDrawer.Height = this.ClientSize.Height;
            pnlDrawerScroll.Height = Math.Max(100, pnlNavDrawer.Height - 48 - 76);

            pnlNavDrawer.Visible = true;
            pnlNavDrawer.BringToFront();
        }

        private void CloseNavMenu()
        {
            menuOpen = false;
            pnlNavDrawer.Visible = false;
            pnlMenuOverlay.Visible = false;
        }

        private void SelectCalculatorMode(string modeName)
        {
            currentMode = modeName;
            CloseNavMenu();

            // Always keep "Created by LuizinhoSouzaK" on the main header title!
            lblTitle.Text = "Created by LuizinhoSouzaK";
            lblModeSubtitle.Text = isAlwaysOnTop ? (currentMode + " • 📌 Sobrepondo Apps") : currentMode;

            // Configure special panels if Converter or Date
            if (converterUnits.ContainsKey(currentMode))
            {
                pnlModeExtra.Visible = true;
                cmbUnitFrom.Visible = true;
                cmbUnitTo.Visible = true;
                lblConvertedPreview.Visible = true;
                dtpStart.Visible = false;
                dtpEnd.Visible = false;
                lblDateDiffResult.Visible = false;

                cmbUnitFrom.Items.Clear();
                cmbUnitTo.Items.Clear();
                foreach (var u in converterUnits[currentMode])
                {
                    cmbUnitFrom.Items.Add(u);
                    cmbUnitTo.Items.Add(u);
                }
                cmbUnitFrom.SelectedIndex = 0;
                cmbUnitTo.SelectedIndex = cmbUnitTo.Items.Count > 1 ? 1 : 0;
                UpdateConversionPreview();
            }
            else if (currentMode == "Cálculo de Data")
            {
                pnlModeExtra.Visible = true;
                cmbUnitFrom.Visible = false;
                cmbUnitTo.Visible = false;
                lblConvertedPreview.Visible = false;
                dtpStart.Visible = true;
                dtpEnd.Visible = true;
                lblDateDiffResult.Visible = true;
                UpdateDateCalculation();
            }
            else
            {
                pnlModeExtra.Visible = false;
            }

            BuildKeypadForCurrentMode();
            RefreshDisplay();
            PerformResponsiveLayout();
        }

        private void UpdateConversionPreview()
        {
            if (!converterUnits.ContainsKey(currentMode)) return;
            UnitDef fromU = cmbUnitFrom.SelectedItem as UnitDef;
            UnitDef toU = cmbUnitTo.SelectedItem as UnitDef;
            if (fromU == null || toU == null) return;

            double inputVal = GetDisplayValue();
            double converted = ConvertValue(currentMode, inputVal, fromU, toU);
            lblConvertedPreview.Text = string.Format("{0} {1} = {2} {3}",
                FormatNumber(inputVal), fromU.Symbol, FormatNumber(converted), toU.Symbol);
        }

        private double ConvertValue(string category, double val, UnitDef fromU, UnitDef toU)
        {
            if (category == "Temperatura")
            {
                double celsius = val;
                if (fromU.Symbol == "°F") celsius = (val - 32.0) * 5.0 / 9.0;
                else if (fromU.Symbol == "K") celsius = val - 273.15;

                if (toU.Symbol == "°C") return celsius;
                if (toU.Symbol == "°F") return (celsius * 9.0 / 5.0) + 32.0;
                if (toU.Symbol == "K") return celsius + 273.15;
                return celsius;
            }
            else
            {
                double baseVal = val * fromU.FactorToBase;
                return baseVal / toU.FactorToBase;
            }
        }

        private void UpdateDateCalculation()
        {
            DateTime d1 = dtpStart.Value.Date;
            DateTime d2 = dtpEnd.Value.Date;
            TimeSpan diff = d2 - d1;
            int totalDays = Math.Abs((int)diff.TotalDays);
            int weeks = totalDays / 7;
            int remDays = totalDays % 7;

            if (totalDays == 0)
            {
                lblDateDiffResult.Text = "Mesmas datas (0 dias)";
            }
            else
            {
                lblDateDiffResult.Text = string.Format("Diferença: {0} dia(s)  ({1} sem, {2} d)", totalDays, weeks, remDays);
            }
            SetDisplayValue(totalDays);
            lblExpression.Text = string.Format("De {0:dd/MM/yyyy} até {1:dd/MM/yyyy}", d1, d2);
        }

        private void BuildKeypadForCurrentMode()
        {
            pnlKeypad.SuspendLayout();
            pnlKeypad.Controls.Clear();
            keypadButtons.Clear();

            string[,] gridLabels;
            if (currentMode == "Científica" || currentMode == "Representação Gráfica")
            {
                gridLabels = new string[7, 5]
                {
                    { degMode ? "DEG" : "RAD", "sin", "cos", "tan", "⌫" },
                    { "xʸ",  "log", "ln",  "(",   ")" },
                    { "²√x", "C",   "CE",  "%",   "÷" },
                    { "x!",  "7",   "8",   "9",   "×" },
                    { "¹/x", "4",   "5",   "6",   "−" },
                    { "π",   "1",   "2",   "3",   "+" },
                    { "x²",  "+/−", "0",   ",",   "=" }
                };
            }
            else if (currentMode == "Programador")
            {
                gridLabels = new string[6, 5]
                {
                    { "A",  "AND", "OR",  "C",   "⌫" },
                    { "B",  "XOR", "NOT", "CE",  "÷" },
                    { "C_HEX", "7", "8",  "9",   "×" },
                    { "D",  "4",   "5",   "6",   "−" },
                    { "E",  "1",   "2",   "3",   "+" },
                    { "F",  "+/−", "0",   ",",   "=" }
                };
            }
            else
            {
                // Padrão, Cálculo de Data & Conversores (6 rows x 4 cols)
                gridLabels = new string[6, 4]
                {
                    { "%",   "CE",  "C",   "⌫" },
                    { "¹/x", "x²",  "²√x", "÷" },
                    { "7",   "8",   "9",   "×" },
                    { "4",   "5",   "6",   "−" },
                    { "1",   "2",   "3",   "+" },
                    { "+/−", "0",   ",",   "=" }
                };
            }

            int rows = gridLabels.GetLength(0);
            int cols = gridLabels.GetLength(1);

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    string rawToken = gridLabels[r, c];
                    string displayLabel = rawToken == "C_HEX" ? "C (12)" : rawToken;

                    bool isDigit = (rawToken.Length == 1 && char.IsDigit(rawToken[0])) || rawToken == "," || rawToken == "+/−";
                    bool isEquals = (rawToken == "=");

                    Button btn = new Button
                    {
                        Text = displayLabel,
                        Tag = new Point(c, r), // store grid coordinate for responsive layout
                        FlatStyle = FlatStyle.Flat,
                        TabStop = false,
                        Cursor = Cursors.Hand
                    };
                    btn.FlatAppearance.BorderSize = 0;

                    if (isEquals)
                    {
                        btn.BackColor = BtnEqualBg;
                        btn.ForeColor = Color.FromArgb(24, 24, 24);
                        btn.FlatAppearance.MouseOverBackColor = BtnEqualHover;
                    }
                    else if (isDigit)
                    {
                        btn.BackColor = BtnNumBg;
                        btn.ForeColor = TextWhite;
                        btn.FlatAppearance.MouseOverBackColor = BtnFuncBg;
                    }
                    else
                    {
                        btn.BackColor = BtnFuncBg;
                        btn.ForeColor = TextWhite;
                        btn.FlatAppearance.MouseOverBackColor = BtnHoverBg;
                    }

                    string captured = rawToken;
                    btn.Click += (s, e) =>
                    {
                        if (menuOpen) CloseNavMenu();
                        OnCalculatorButtonClick(captured);
                        this.ActiveControl = null;
                    };

                    keypadButtons.Add(btn);
                    pnlKeypad.Controls.Add(btn);
                }
            }

            pnlKeypad.ResumeLayout();
        }

        // --- Fully Responsive Layout Engine ---
        private void PerformResponsiveLayout()
        {
            if (this.WindowState == FormWindowState.Minimized || pnlKeypad == null)
                return;

            this.SuspendLayout();

            int totalW = this.ClientSize.Width;
            int totalH = this.ClientSize.Height;

            // Auto-dock History side panel when window is wide (>= 650px) or when manually toggled
            historyAutoDocked = (totalW >= 650);
            bool showHistory = historyUserToggled || historyAutoDocked;
            int historyW = 0;

            if (showHistory)
            {
                historyW = Math.Max(250, Math.Min(380, (int)(totalW * 0.42)));
                if (totalW - historyW < 270)
                {
                    historyW = Math.Max(0, totalW - 270);
                }
            }

            int calcW = Math.Max(260, totalW - historyW);
            int pad = Math.Max(6, Math.Min(12, calcW / 45));

            // 1. Header Bar
            int headerH = 42;
            btnMenu.Location = new Point(pad, 6);
            btnMenu.Size = new Size(36, 34);

            btnToggleHistory.Size = new Size(36, 34);
            btnToggleHistory.Location = new Point(calcW - pad - 36, 6);

            btnAlwaysOnTop.Size = new Size(36, 34);
            btnAlwaysOnTop.Location = new Point(btnToggleHistory.Left - 38, 6);

            int titleLeft = btnMenu.Right + 6;
            int titleWidth = Math.Max(80, btnAlwaysOnTop.Left - titleLeft - 4);
            lblTitle.Location = new Point(titleLeft, 5);
            lblTitle.Size = new Size(titleWidth, 20);
            float titleFontSz = calcW < 310 ? 9.5f : 11f;
            lblTitle.Font = new Font("Segoe UI Semibold", titleFontSz, FontStyle.Bold);

            lblModeSubtitle.Location = new Point(titleLeft, 24);
            lblModeSubtitle.Size = new Size(titleWidth, 16);

            // 2. Expression & Display Area
            int curY = headerH + 2;
            int exprH = Math.Max(20, (int)(totalH * 0.042));
            lblExpression.Location = new Point(pad, curY);
            lblExpression.Size = new Size(calcW - pad * 2, exprH);
            curY += exprH;

            int dispH = Math.Max(46, Math.Min(110, (int)(totalH * 0.115)));
            lblDisplay.Location = new Point(pad, curY);
            lblDisplay.Size = new Size(calcW - pad * 2, dispH);
            UpdateResponsiveDisplayFont();
            curY += dispH + 2;

            // 2b. Optional Mode Panel (Converter / Date)
            if (pnlModeExtra.Visible)
            {
                int extraH = 62;
                pnlModeExtra.Location = new Point(pad, curY);
                pnlModeExtra.Size = new Size(calcW - pad * 2, extraH);

                int innerW = pnlModeExtra.Width - 12;
                int halfW = (innerW - 8) / 2;

                if (cmbUnitFrom.Visible)
                {
                    cmbUnitFrom.Location = new Point(6, 6);
                    cmbUnitFrom.Size = new Size(halfW, 24);
                    cmbUnitTo.Location = new Point(6 + halfW + 8, 6);
                    cmbUnitTo.Size = new Size(halfW, 24);
                    lblConvertedPreview.Location = new Point(6, 34);
                    lblConvertedPreview.Size = new Size(innerW, 24);
                }
                else if (dtpStart.Visible)
                {
                    dtpStart.Location = new Point(6, 6);
                    dtpStart.Size = new Size(halfW, 24);
                    dtpEnd.Location = new Point(6 + halfW + 8, 6);
                    dtpEnd.Size = new Size(halfW, 24);
                    lblDateDiffResult.Location = new Point(6, 34);
                    lblDateDiffResult.Size = new Size(innerW, 24);
                }

                curY += extraH + 4;
            }

            // 3. Memory Row (6 buttons distributed evenly across calcW)
            int memH = Math.Max(24, Math.Min(36, (int)(totalH * 0.048)));
            int availW = calcW - pad * 2;
            int memBtnW = availW / 6;
            for (int i = 0; i < 6; i++)
            {
                memButtons[i].Location = new Point(pad + i * memBtnW, curY);
                memButtons[i].Size = new Size(memBtnW - 2, memH);
            }
            curY += memH + 4;

            // 4. Print Report & Clear History Bar
            int actionH = Math.Max(28, Math.Min(40, (int)(totalH * 0.055)));
            int printW = (int)(availW * 0.55) - 2;
            int clearW = availW - printW - 4;
            btnPrintReport.Location = new Point(pad, curY);
            btnPrintReport.Size = new Size(printW, actionH);

            btnClearHistory.Location = new Point(pad + printW + 4, curY);
            btnClearHistory.Size = new Size(clearW, actionH);
            curY += actionH + 6;

            // 5. Keypad Grid Area (fills remaining height & width of calculator area)
            int keypadH = Math.Max(180, totalH - curY - pad);
            pnlKeypad.Location = new Point(pad, curY);
            pnlKeypad.Size = new Size(availW, keypadH);

            if (keypadButtons.Count > 0)
            {
                int cols = (currentMode == "Científica" || currentMode == "Representação Gráfica" || currentMode == "Programador") ? 5 : 4;
                int rows = (currentMode == "Científica" || currentMode == "Representação Gráfica") ? 7 : 6;
                int gap = Math.Max(2, Math.Min(5, availW / 110));

                float cellW = (float)(availW - gap * (cols - 1)) / cols;
                float cellH = (float)(keypadH - gap * (rows - 1)) / rows;

                float baseFontSz = Math.Max(9.5f, Math.Min(20f, Math.Min(cellW * 0.23f, cellH * 0.32f)));

                foreach (Button btn in keypadButtons)
                {
                    Point gridPos = (Point)btn.Tag;
                    int c = gridPos.X;
                    int r = gridPos.Y;

                    int x = (int)Math.Round(c * (cellW + gap));
                    int y = (int)Math.Round(r * (cellH + gap));
                    int nextX = (int)Math.Round((c + 1) * (cellW + gap)) - gap;
                    int nextY = (int)Math.Round((r + 1) * (cellH + gap)) - gap;

                    btn.Location = new Point(x, y);
                    btn.Size = new Size(Math.Max(20, nextX - x), Math.Max(20, nextY - y));

                    bool isEquals = btn.Text == "=";
                    bool isNum = (btn.Text.Length == 1 && char.IsDigit(btn.Text[0])) || btn.Text == "," || btn.Text == "+/−";
                    if (isEquals)
                        btn.Font = new Font("Segoe UI Semibold", baseFontSz * 1.15f, FontStyle.Bold);
                    else if (isNum)
                        btn.Font = new Font("Segoe UI Semibold", baseFontSz * 1.05f, FontStyle.Bold);
                    else
                        btn.Font = new Font("Segoe UI", baseFontSz * 0.92f, FontStyle.Regular);
                }
            }

            // 6. Side History Panel
            pnlHistorySide.Visible = showHistory && historyW > 100;
            if (pnlHistorySide.Visible)
            {
                pnlHistorySide.Location = new Point(calcW, 0);
                pnlHistorySide.Size = new Size(historyW, totalH);

                int hPad = 12;
                int hInnerW = Math.Max(100, historyW - hPad * 2);
                lblHistHeader.Location = new Point(hPad, 10);
                lblHistHeader.Size = new Size(hInnerW, 24);

                lblHistoryCount.Location = new Point(hPad, 34);
                lblHistoryCount.Size = new Size(hInnerW, 18);

                int bottomBtnH = 40;
                int listH = Math.Max(80, totalH - 58 - bottomBtnH - 20);
                lstHistory.Location = new Point(hPad, 56);
                lstHistory.Size = new Size(hInnerW, listH);

                int sideBtnW = (hInnerW - 8) / 2;
                int sideBtnY = lstHistory.Bottom + 8;
                btnSidePrint.Location = new Point(hPad, sideBtnY);
                btnSidePrint.Size = new Size(sideBtnW, bottomBtnH);

                btnSideClear.Location = new Point(hPad + sideBtnW + 8, sideBtnY);
                btnSideClear.Size = new Size(sideBtnW, bottomBtnH);
            }

            // 7. Navigation Drawer Height if open
            if (menuOpen && pnlNavDrawer != null)
            {
                pnlNavDrawer.Height = totalH;
                pnlDrawerScroll.Height = Math.Max(100, totalH - 48 - 76);
                pnlNavDrawer.BringToFront();
            }

            this.ResumeLayout();
        }

        private void UpdateResponsiveDisplayFont()
        {
            if (lblDisplay == null) return;
            string text = lblDisplay.Text ?? "0";
            int len = text.Length;

            float maxSz = Math.Max(18f, Math.Min(44f, lblDisplay.Height * 0.52f));
            float widthFactor = Math.Max(12f, (lblDisplay.Width / Math.Max(8f, len * 0.68f)));
            float finalSz = Math.Max(13f, Math.Min(maxSz, widthFactor));

            lblDisplay.Font = new Font("Segoe UI Semibold", finalSz, FontStyle.Bold);
        }

        private Button CreateFlatIconButton(string text, Point location, Size size, float fontSize)
        {
            Button btn = new Button
            {
                Text = text,
                Font = new Font("Segoe UI Symbol", fontSize, FontStyle.Regular),
                ForeColor = TextWhite,
                BackColor = BgDark,
                FlatStyle = FlatStyle.Flat,
                Location = location,
                Size = size,
                TabStop = false,
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = BtnFuncBg;
            return btn;
        }

        private void ToggleHistoryPanel()
        {
            if (historyAutoDocked)
            {
                // If window is already wide enough to auto-dock, shrink/expand or toggle user state
                historyUserToggled = !historyUserToggled;
                if (!historyUserToggled)
                {
                    this.ClientSize = new Size(360, this.ClientSize.Height);
                }
            }
            else
            {
                historyUserToggled = !historyUserToggled;
                if (historyUserToggled && this.ClientSize.Width < 650)
                {
                    this.ClientSize = new Size(660, this.ClientSize.Height);
                }
                else if (!historyUserToggled && this.ClientSize.Width >= 650 && this.WindowState == FormWindowState.Normal)
                {
                    this.ClientSize = new Size(360, this.ClientSize.Height);
                }
            }
            PerformResponsiveLayout();
        }

        // --- Calculator Logic ---
        private void OnCalculatorButtonClick(string token)
        {
            if (token.Length == 1 && char.IsDigit(token[0]))
            {
                AppendDigit(token);
                return;
            }

            switch (token)
            {
                case "A": AppendDigit("10"); break;
                case "B": AppendDigit("11"); break;
                case "C_HEX": AppendDigit("12"); break;
                case "D": AppendDigit("13"); break;
                case "E": AppendDigit("14"); break;
                case "F": AppendDigit("15"); break;
                case ",":
                    AppendDecimal();
                    break;
                case "+/−":
                    ToggleSign();
                    break;
                case "⌫":
                    Backspace();
                    break;
                case "CE":
                    currentInput = "0";
                    startNewInput = true;
                    RefreshDisplay();
                    break;
                case "C":
                    currentInput = "0";
                    accumulator = 0;
                    pendingOp = "";
                    lastBinaryOp = "";
                    lastOperand = 0;
                    startNewInput = true;
                    lastWasEquals = false;
                    lblExpression.Text = "";
                    RefreshDisplay();
                    break;
                case "%":
                    ApplyPercentage();
                    break;
                case "π":
                    SetDisplayValue(Math.PI);
                    startNewInput = true;
                    break;
                case "DEG":
                case "RAD":
                    degMode = !degMode;
                    BuildKeypadForCurrentMode();
                    PerformResponsiveLayout();
                    break;
                case "sin":
                    ApplyUnaryOperation("sin(", ")", x => Math.Sin(degMode ? x * Math.PI / 180.0 : x));
                    break;
                case "cos":
                    ApplyUnaryOperation("cos(", ")", x => Math.Cos(degMode ? x * Math.PI / 180.0 : x));
                    break;
                case "tan":
                    ApplyUnaryOperation("tan(", ")", x => Math.Tan(degMode ? x * Math.PI / 180.0 : x));
                    break;
                case "log":
                    ApplyUnaryOperation("log(", ")", x =>
                    {
                        if (x <= 0) throw new ArithmeticException("Entrada inválida");
                        return Math.Log10(x);
                    });
                    break;
                case "ln":
                    ApplyUnaryOperation("ln(", ")", x =>
                    {
                        if (x <= 0) throw new ArithmeticException("Entrada inválida");
                        return Math.Log(x);
                    });
                    break;
                case "x!":
                    ApplyUnaryOperation("fact(", ")", x =>
                    {
                        if (x < 0 || x > 170 || Math.Floor(x) != x) throw new ArithmeticException("Entrada inválida");
                        double r = 1;
                        for (int i = 2; i <= (int)x; i++) r *= i;
                        return r;
                    });
                    break;
                case "¹/x":
                    ApplyUnaryOperation("1/(", ")", x =>
                    {
                        if (x == 0) throw new DivideByZeroException();
                        return 1.0 / x;
                    });
                    break;
                case "x²":
                    ApplyUnaryOperation("sqr(", ")", x => x * x);
                    break;
                case "²√x":
                    ApplyUnaryOperation("√(", ")", x =>
                    {
                        if (x < 0) throw new ArithmeticException("Entrada inválida");
                        return Math.Sqrt(x);
                    });
                    break;
                case "NOT":
                    ApplyUnaryOperation("NOT(", ")", x => (double)(~(long)x));
                    break;
                case "+":
                case "−":
                case "×":
                case "÷":
                case "xʸ":
                case "AND":
                case "OR":
                case "XOR":
                    ApplyBinaryOperator(token);
                    break;
                case "=":
                    if (converterUnits.ContainsKey(currentMode))
                    {
                        UnitDef fromU = cmbUnitFrom.SelectedItem as UnitDef;
                        UnitDef toU = cmbUnitTo.SelectedItem as UnitDef;
                        if (fromU != null && toU != null)
                        {
                            double v = GetDisplayValue();
                            double conv = ConvertValue(currentMode, v, fromU, toU);
                            string expr = string.Format("[{0}] {1} {2} → {3} =", currentMode, FormatNumber(v), fromU.Symbol, toU.Symbol);
                            lblExpression.Text = expr;
                            RecordHistory(expr, conv);
                            UpdateConversionPreview();
                            return;
                        }
                    }
                    EvaluateEquals();
                    break;
            }
        }

        private void AppendDigit(string digit)
        {
            if (lastWasEquals)
            {
                lblExpression.Text = "";
                pendingOp = "";
                lastWasEquals = false;
            }

            if (startNewInput || currentInput == "0")
            {
                currentInput = digit;
                startNewInput = false;
            }
            else
            {
                string digitsOnly = currentInput.Replace("-", "").Replace(",", "");
                if (digitsOnly.Length < 15)
                {
                    currentInput += digit;
                }
            }
            RefreshDisplay();
        }

        private void AppendDecimal()
        {
            if (lastWasEquals)
            {
                lblExpression.Text = "";
                pendingOp = "";
                lastWasEquals = false;
            }

            if (startNewInput)
            {
                currentInput = "0,";
                startNewInput = false;
            }
            else if (!currentInput.Contains(","))
            {
                currentInput += ",";
            }
            RefreshDisplay();
        }

        private void ToggleSign()
        {
            if (currentInput == "0" || string.IsNullOrEmpty(currentInput))
                return;

            if (currentInput.StartsWith("-"))
                currentInput = currentInput.Substring(1);
            else
                currentInput = "-" + currentInput;

            RefreshDisplay();
        }

        private void Backspace()
        {
            if (startNewInput || lastWasEquals)
                return;

            if (currentInput.Length > 1)
            {
                currentInput = currentInput.Substring(0, currentInput.Length - 1);
                if (currentInput == "-" || currentInput == "-0")
                    currentInput = "0";
            }
            else
            {
                currentInput = "0";
            }
            RefreshDisplay();
        }

        private void ApplyPercentage()
        {
            double val = GetDisplayValue();
            double result;
            string expr;

            if (!string.IsNullOrEmpty(pendingOp))
            {
                if (pendingOp == "+" || pendingOp == "−")
                    result = accumulator * (val / 100.0);
                else
                    result = val / 100.0;

                expr = string.Format("{0} {1} {2}% =", FormatNumber(accumulator), pendingOp, FormatNumber(val));
            }
            else
            {
                result = val / 100.0;
                expr = string.Format("{0}% =", FormatNumber(val));
                RecordHistory(expr, result);
            }

            SetDisplayValue(result);
            startNewInput = true;
        }

        private void ApplyUnaryOperation(string prefix, string suffix, Func<double, double> op)
        {
            double val = GetDisplayValue();
            try
            {
                double result = op(val);
                string expr = string.Format("{0}{1}{2} =", prefix, FormatNumber(val), suffix);
                lblExpression.Text = expr;
                SetDisplayValue(result);
                RecordHistory(expr, result);
                startNewInput = true;
                lastWasEquals = true;
            }
            catch (DivideByZeroException)
            {
                lblDisplay.Text = "Não é possível dividir por zero";
                UpdateResponsiveDisplayFont();
                startNewInput = true;
            }
            catch (Exception)
            {
                lblDisplay.Text = "Entrada inválida";
                UpdateResponsiveDisplayFont();
                startNewInput = true;
            }
        }

        private void ApplyBinaryOperator(string op)
        {
            double currentVal = GetDisplayValue();

            if (!string.IsNullOrEmpty(pendingOp) && !startNewInput)
            {
                double intermediate;
                if (!TryComputeBinary(accumulator, currentVal, pendingOp, out intermediate))
                    return;

                string stepExpr = string.Format("{0} {1} {2} =", FormatNumber(accumulator), pendingOp, FormatNumber(currentVal));
                RecordHistory(stepExpr, intermediate);
                accumulator = intermediate;
                SetDisplayValue(accumulator);
            }
            else
            {
                accumulator = currentVal;
            }

            pendingOp = op;
            lastBinaryOp = op;
            lastWasEquals = false;
            startNewInput = true;
            lblExpression.Text = string.Format("{0} {1}", FormatNumber(accumulator), pendingOp);
        }

        private void EvaluateEquals()
        {
            double currentVal = GetDisplayValue();

            if (!string.IsNullOrEmpty(pendingOp))
            {
                lastOperand = currentVal;
                lastBinaryOp = pendingOp;
                string expr = string.Format("{0} {1} {2} =", FormatNumber(accumulator), pendingOp, FormatNumber(currentVal));

                double res;
                if (!TryComputeBinary(accumulator, currentVal, pendingOp, out res))
                    return;

                lblExpression.Text = expr;
                SetDisplayValue(res);
                RecordHistory(expr, res);

                accumulator = res;
                pendingOp = "";
                startNewInput = true;
                lastWasEquals = true;
            }
            else if (!string.IsNullOrEmpty(lastBinaryOp))
            {
                string expr = string.Format("{0} {1} {2} =", FormatNumber(currentVal), lastBinaryOp, FormatNumber(lastOperand));
                double res;
                if (!TryComputeBinary(currentVal, lastOperand, lastBinaryOp, out res))
                    return;

                lblExpression.Text = expr;
                SetDisplayValue(res);
                RecordHistory(expr, res);

                accumulator = res;
                startNewInput = true;
                lastWasEquals = true;
            }
        }

        private bool TryComputeBinary(double left, double right, string op, out double result)
        {
            result = 0;
            switch (op)
            {
                case "+":
                    result = left + right;
                    return true;
                case "−":
                    result = left - right;
                    return true;
                case "×":
                    result = left * right;
                    return true;
                case "xʸ":
                    result = Math.Pow(left, right);
                    return true;
                case "AND":
                    result = (double)((long)left & (long)right);
                    return true;
                case "OR":
                    result = (double)((long)left | (long)right);
                    return true;
                case "XOR":
                    result = (double)((long)left ^ (long)right);
                    return true;
                case "÷":
                    if (right == 0)
                    {
                        lblDisplay.Text = "Não é possível dividir por zero";
                        UpdateResponsiveDisplayFont();
                        pendingOp = "";
                        startNewInput = true;
                        return false;
                    }
                    result = left / right;
                    return true;
            }
            return false;
        }

        private double GetDisplayValue()
        {
            double v;
            if (double.TryParse(currentInput, NumberStyles.Any, ptBr, out v))
                return v;
            return 0;
        }

        private void SetDisplayValue(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                currentInput = "0";
                lblDisplay.Text = "Resultado indefinido";
                UpdateResponsiveDisplayFont();
                return;
            }
            currentInput = value.ToString("G15", ptBr);
            RefreshDisplay();
        }

        private string FormatNumber(double value)
        {
            return value.ToString("#,##0.###############", ptBr);
        }

        private void RefreshDisplay()
        {
            string formatted = currentInput;
            double val;
            if (double.TryParse(currentInput, NumberStyles.Any, ptBr, out val))
            {
                string[] parts = currentInput.Split(',');
                double intPart;
                if (double.TryParse(parts[0], NumberStyles.Any, ptBr, out intPart))
                {
                    string sign = (currentInput.StartsWith("-") && intPart == 0) ? "-" : "";
                    formatted = sign + intPart.ToString("#,##0", ptBr);
                    if (parts.Length > 1)
                        formatted += "," + parts[1];
                }

                if (currentMode == "Programador")
                {
                    long intVal = (long)val;
                    lblExpression.Text = string.Format("HEX: {0:X}   OCT: {1}   BIN: {2}",
                        intVal, Convert.ToString(intVal, 8), Convert.ToString(intVal, 2));
                }
            }

            lblDisplay.Text = formatted;
            UpdateResponsiveDisplayFont();

            if (converterUnits.ContainsKey(currentMode))
            {
                UpdateConversionPreview();
            }
        }

        // --- Memory Functions ---
        private void ModifyMemory(double val, bool add)
        {
            if (memoryStack.Count == 0)
                memoryStack.Add(add ? val : -val);
            else
                memoryStack[0] += (add ? val : -val);

            UpdateMemoryButtonsState();
            startNewInput = true;
        }

        private void UpdateMemoryButtonsState()
        {
            bool hasMem = memoryStack.Count > 0;
            btnMc.Enabled = hasMem;
            btnMr.Enabled = hasMem;
            btnMv.Enabled = hasMem;
            btnMc.ForeColor = hasMem ? TextWhite : TextDisabled;
            btnMr.ForeColor = hasMem ? TextWhite : TextDisabled;
            btnMv.ForeColor = hasMem ? TextWhite : TextDisabled;
        }

        private void ShowMemoryPopup()
        {
            if (memoryStack.Count == 0) return;
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Valores armazenados no registrador M:");
            sb.AppendLine("--------------------------------------");
            for (int i = 0; i < memoryStack.Count; i++)
            {
                sb.AppendFormat("M[{0}] = {1}\r\n", i + 1, FormatNumber(memoryStack[i]));
            }
            MessageBox.Show(sb.ToString(), "Memória (M▾) — Created by LuizinhoSouzaK", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowAboutDialog()
        {
            MessageBox.Show(
                "Calculadora do Windows Customizada\n" +
                "Created by LuizinhoSouzaK\n\n" +
                "• Menu Opções completo idêntico ao da Calculadora original (Padrão, Científica, Programador, Cálculo de Data e 13 Conversores)\n" +
                "• Botão 'Manter no Topo' (Sobrepor outros aplicativos) integrado no cabeçalho\n" +
                "• Layout 100% responsivo (redimensionável e tela cheia)\n" +
                "• Armazena todos os cálculos na memória automaticamente\n" +
                "• Emissão e impressão de relatório completo (Impressora/PDF/HTML)",
                "Sobre — Created by LuizinhoSouzaK",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        // --- Calculation History & Printing ---
        private void RecordHistory(string expression, double result)
        {
            var rec = new CalculationRecord
            {
                Index = calculationHistory.Count + 1,
                Timestamp = DateTime.Now,
                Expression = expression,
                ResultValue = result,
                ResultFormatted = FormatNumber(result)
            };

            calculationHistory.Add(rec);
            lstHistory.Items.Add(rec.ToString());
            lstHistory.TopIndex = lstHistory.Items.Count - 1;

            SaveHistoryToDisk();
            UpdateHistoryBadge();
        }

        private void UpdateHistoryBadge()
        {
            int count = calculationHistory.Count;
            btnPrintReport.Text = string.Format("🖨 Imprimir Relatório ({0})", count);
            lblHistoryCount.Text = count == 0
                ? "Nenhum cálculo registrado na memória."
                : string.Format("{0} cálculo(s) na memória (duplo-clique p/ usar)", count);
        }

        private void SaveHistoryToDisk()
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                foreach (var r in calculationHistory)
                {
                    sb.AppendFormat("{0:yyyy-MM-dd HH:mm:ss}\t{1}\t{2}\t{3}\r\n",
                        r.Timestamp, r.Expression, r.ResultFormatted, r.ResultValue.ToString("R", CultureInfo.InvariantCulture));
                }
                File.WriteAllText(historyFilePath, sb.ToString(), Encoding.UTF8);
            }
            catch { }
        }

        private void LoadSavedHistory()
        {
            try
            {
                if (!File.Exists(historyFilePath)) return;
                string[] lines = File.ReadAllLines(historyFilePath, Encoding.UTF8);
                foreach (string line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    string[] parts = line.Split('\t');
                    if (parts.Length >= 4)
                    {
                        DateTime parsedDt;
                        DateTime dt = DateTime.TryParse(parts[0], out parsedDt) ? parsedDt : DateTime.Now;
                        double parsedVal;
                        double val = double.TryParse(parts[3], NumberStyles.Any, CultureInfo.InvariantCulture, out parsedVal) ? parsedVal : 0;
                        var rec = new CalculationRecord
                        {
                            Index = calculationHistory.Count + 1,
                            Timestamp = dt,
                            Expression = parts[1],
                            ResultFormatted = parts[2],
                            ResultValue = val
                        };
                        calculationHistory.Add(rec);
                        lstHistory.Items.Add(rec.ToString());
                    }
                }
            }
            catch { }
        }

        private void ClearHistoryAction()
        {
            if (calculationHistory.Count == 0)
            {
                MessageBox.Show(
                    "O histórico de cálculos na memória já está vazio!",
                    "Limpar Histórico — Created by LuizinhoSouzaK",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            DialogResult dr = MessageBox.Show(
                string.Format("Deseja realmente limpar todos os {0} cálculo(s) armazenados no histórico da memória?", calculationHistory.Count),
                "Confirmar Limpeza de Histórico",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (dr == DialogResult.Yes)
            {
                calculationHistory.Clear();
                lstHistory.Items.Clear();
                try
                {
                    if (File.Exists(historyFilePath))
                        File.Delete(historyFilePath);
                }
                catch { }
                UpdateHistoryBadge();
                MessageBox.Show(
                    "Histórico de cálculos limpo com sucesso!",
                    "Histórico Limpo — Created by LuizinhoSouzaK",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private void ShowPrintReportOptions()
        {
            if (calculationHistory.Count == 0)
            {
                MessageBox.Show(
                    "Ainda não há cálculos armazenados na memória.\nRealize algumas operações antes de imprimir o relatório!",
                    "Relatório de Histórico — Created by LuizinhoSouzaK",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            using (Form dlg = new Form())
            {
                dlg.Text = "Imprimir Relatório do Histórico";
                dlg.BackColor = BgDark;
                dlg.ForeColor = TextWhite;
                dlg.Font = new Font("Segoe UI", 9.5F);
                dlg.ClientSize = new Size(420, 250);
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.MaximizeBox = false;
                dlg.MinimizeBox = false;
                dlg.TopMost = this.TopMost;

                Label lblInfo = new Label
                {
                    Text = string.Format("Total de cálculos armazenados na memória: {0}\nEscolha como deseja emitir o relatório:", calculationHistory.Count),
                    Location = new Point(16, 16),
                    Size = new Size(388, 44),
                    ForeColor = TextWhite
                };

                Button btnDirectPrint = new Button
                {
                    Text = "🖨️ Visualizar e Imprimir na Impressora / PDF (Windows)",
                    Location = new Point(16, 70),
                    Size = new Size(388, 44),
                    BackColor = Color.FromArgb(42, 68, 92),
                    ForeColor = TextWhite,
                    FlatStyle = FlatStyle.Flat,
                    Cursor = Cursors.Hand
                };
                btnDirectPrint.FlatAppearance.BorderColor = AccentBlue;
                btnDirectPrint.Click += (s, e) =>
                {
                    dlg.Close();
                    PrintToWindowsPrinter();
                };

                Button btnHtmlReport = new Button
                {
                    Text = "📄 Gerar Relatório HTML Formatado (Abre no Navegador)",
                    Location = new Point(16, 124),
                    Size = new Size(388, 44),
                    BackColor = BtnNumBg,
                    ForeColor = TextWhite,
                    FlatStyle = FlatStyle.Flat,
                    Cursor = Cursors.Hand
                };
                btnHtmlReport.Click += (s, e) =>
                {
                    dlg.Close();
                    ExportAndOpenHtmlReport();
                };

                Button btnViewHistory = new Button
                {
                    Text = "📜 Ver Painel Lateral de Histórico na Calculadora",
                    Location = new Point(16, 178),
                    Size = new Size(388, 44),
                    BackColor = BtnFuncBg,
                    ForeColor = TextWhite,
                    FlatStyle = FlatStyle.Flat,
                    Cursor = Cursors.Hand
                };
                btnViewHistory.Click += (s, e) =>
                {
                    dlg.Close();
                    if (!pnlHistorySide.Visible) ToggleHistoryPanel();
                };

                dlg.Controls.Add(lblInfo);
                dlg.Controls.Add(btnDirectPrint);
                dlg.Controls.Add(btnHtmlReport);
                dlg.Controls.Add(btnViewHistory);
                dlg.ShowDialog(this);
            }
        }

        private void PrintToWindowsPrinter()
        {
            PrintDocument pd = new PrintDocument();
            pd.DocumentName = "Relatorio_Calculos_LuizinhoSouzaK";
            printCurrentIndex = 0;

            pd.PrintPage += (sender, e) =>
            {
                Graphics g = e.Graphics;
                Font titleFont = new Font("Segoe UI", 15F, FontStyle.Bold);
                Font subFont = new Font("Segoe UI", 10F, FontStyle.Italic);
                Font headerFont = new Font("Consolas", 10F, FontStyle.Bold);
                Font rowFont = new Font("Consolas", 10F, FontStyle.Regular);

                float x = e.MarginBounds.Left;
                float y = e.MarginBounds.Top;
                float width = e.MarginBounds.Width;

                g.DrawString("RELATÓRIO DE HISTÓRICO DE CÁLCULOS", titleFont, Brushes.Black, x, y);
                y += 28;
                g.DrawString("Calculadora — Created by LuizinhoSouzaK", subFont, Brushes.DarkSlateGray, x, y);
                y += 20;
                g.DrawString(string.Format("Emitido em: {0:dd/MM/yyyy HH:mm:ss}   |   Total de Operações: {1}", DateTime.Now, calculationHistory.Count), subFont, Brushes.Black, x, y);
                y += 26;

                g.DrawLine(Pens.Black, x, y, x + width, y);
                y += 8;
                g.DrawString(string.Format("{0,-6} {1,-12} {2,-34} {3,16}", "#", "HORÁRIO", "OPERAÇÃO / EXPRESSÃO", "RESULTADO"), headerFont, Brushes.Black, x, y);
                y += 20;
                g.DrawLine(Pens.Black, x, y, x + width, y);
                y += 10;

                double sumResults = 0;
                foreach (var item in calculationHistory) sumResults += item.ResultValue;

                while (printCurrentIndex < calculationHistory.Count)
                {
                    var rec = calculationHistory[printCurrentIndex];
                    string line = string.Format("{0,-6} {1,-12:HH:mm:ss} {2,-34} {3,16}",
                        rec.Index, rec.Timestamp, rec.Expression, rec.ResultFormatted);

                    g.DrawString(line, rowFont, Brushes.Black, x, y);
                    y += 20;
                    printCurrentIndex++;

                    if (y + 60 > e.MarginBounds.Bottom && printCurrentIndex < calculationHistory.Count)
                    {
                        e.HasMorePages = true;
                        return;
                    }
                }

                y += 10;
                g.DrawLine(Pens.Black, x, y, x + width, y);
                y += 10;
                g.DrawString(string.Format("Soma acumulada dos resultados: {0}", FormatNumber(sumResults)), headerFont, Brushes.Black, x, y);
                e.HasMorePages = false;
            };

            using (PrintPreviewDialog preview = new PrintPreviewDialog())
            {
                preview.Document = pd;
                preview.Width = 900;
                preview.Height = 700;
                preview.StartPosition = FormStartPosition.CenterParent;
                preview.Text = "Visualização de Impressão — Relatório de Cálculos (Created by LuizinhoSouzaK)";
                preview.ShowDialog(this);
            }
        }

        private void ExportAndOpenHtmlReport()
        {
            string htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Relatorio_Historico_Calculos.html");
            double sumResults = 0;
            StringBuilder rows = new StringBuilder();

            foreach (var rec in calculationHistory)
            {
                sumResults += rec.ResultValue;
                rows.AppendFormat(
                    "<tr><td>#{0}</td><td>{1:dd/MM/yyyy HH:mm:ss}</td><td class='expr'>{2}</td><td class='res'>{3}</td></tr>\r\n",
                    rec.Index, rec.Timestamp, System.Net.WebUtility.HtmlEncode(rec.Expression), System.Net.WebUtility.HtmlEncode(rec.ResultFormatted));
            }

            string html = string.Format(@"<!DOCTYPE html>
<html lang='pt-BR'>
<head>
<meta charset='UTF-8'>
<title>Relatório de Histórico de Cálculos — Created by LuizinhoSouzaK</title>
<style>
  body {{ font-family: 'Segoe UI', Tahoma, sans-serif; background: #181818; color: #f0f0f0; margin: 32px; }}
  .card {{ max-width: 820px; margin: auto; background: #242424; padding: 28px; border-radius: 10px; box-shadow: 0 6px 24px rgba(0,0,0,0.45); }}
  h1 {{ margin: 0 0 6px 0; color: #76B9ED; font-size: 24px; }}
  .subtitle {{ color: #bbbbbb; font-size: 14px; margin-bottom: 20px; }}
  .toolbar {{ margin-bottom: 20px; }}
  .btn {{ background: #76B9ED; color: #181818; border: none; padding: 10px 20px; font-weight: bold; border-radius: 6px; cursor: pointer; font-size: 14px; }}
  table {{ width: 100%; border-collapse: collapse; margin-top: 12px; }}
  th, td {{ padding: 10px 12px; border-bottom: 1px solid #383838; text-align: left; }}
  th {{ background: #303030; color: #76B9ED; }}
  td.expr {{ font-family: Consolas, monospace; color: #d0d0d0; }}
  td.res {{ font-family: Consolas, monospace; font-weight: bold; color: #ffffff; text-align: right; font-size: 16px; }}
  .summary {{ margin-top: 20px; padding: 14px; background: #303030; border-left: 4px solid #76B9ED; border-radius: 4px; font-size: 15px; }}
  @media print {{
    body {{ background: #ffffff; color: #000000; margin: 0; }}
    .card {{ box-shadow: none; background: #ffffff; color: #000000; }}
    .toolbar {{ display: none; }}
    h1, th {{ color: #000000; background: #eeeeee; }}
    td.expr, td.res, .subtitle {{ color: #000000; }}
    .summary {{ background: #f5f5f5; border-left: 4px solid #000000; }}
  }}
</style>
</head>
<body>
  <div class='card'>
    <h1>Relatório de Histórico de Cálculos</h1>
    <div class='subtitle'>Calculadora Customizada — <strong>Created by LuizinhoSouzaK</strong><br>Emitido em: {0:dd/MM/yyyy HH:mm:ss} | Total de cálculos armazenados: {1}</div>
    <div class='toolbar'>
      <button class='btn' onclick='window.print()'>🖨️ Imprimir Relatório / Salvar em PDF</button>
    </div>
    <table>
      <thead>
        <tr><th>#</th><th>Data / Hora</th><th>Expressão</th><th style='text-align:right'>Resultado</th></tr>
      </thead>
      <tbody>
        {2}
      </tbody>
    </table>
    <div class='summary'>
      <strong>Soma acumulada dos resultados:</strong> {3}
    </div>
  </div>
</body>
</html>", DateTime.Now, calculationHistory.Count, rows.ToString(), FormatNumber(sumResults));

            File.WriteAllText(htmlPath, html, Encoding.UTF8);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(htmlPath) { UseShellExecute = true });
        }

        // --- Keyboard Shortcuts ---
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            Keys keyCode = keyData & Keys.KeyCode;
            bool alt = (keyData & Keys.Alt) != 0;
            bool ctrl = (keyData & Keys.Control) != 0;

            // Enter (teclado principal ou teclado numérico) SEMPRE executa "=" igual à Calculadora do Windows
            if (!alt && !ctrl && (keyCode == Keys.Enter || keyCode == Keys.Return))
            {
                if (menuOpen) CloseNavMenu();
                OnCalculatorButtonClick("=");
                this.ActiveControl = null;
                return true;
            }

            if (alt && (keyCode == Keys.Up || keyCode == Keys.Down))
            {
                ToggleAlwaysOnTop();
                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override bool ProcessDialogKey(Keys keyData)
        {
            Keys keyCode = keyData & Keys.KeyCode;
            if (keyCode == Keys.Enter || keyCode == Keys.Return)
            {
                if (menuOpen) CloseNavMenu();
                OnCalculatorButtonClick("=");
                this.ActiveControl = null;
                return true;
            }
            return base.ProcessDialogKey(keyData);
        }

        private void MainForm_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return || e.KeyChar == '=')
            {
                if (menuOpen) CloseNavMenu();
                OnCalculatorButtonClick("=");
                e.Handled = true;
            }
            else if (char.IsDigit(e.KeyChar))
            {
                AppendDigit(e.KeyChar.ToString());
                e.Handled = true;
            }
            else if (e.KeyChar == ',' || e.KeyChar == '.')
            {
                AppendDecimal();
                e.Handled = true;
            }
            else if (e.KeyChar == '+')
            {
                ApplyBinaryOperator("+");
                e.Handled = true;
            }
            else if (e.KeyChar == '-')
            {
                ApplyBinaryOperator("−");
                e.Handled = true;
            }
            else if (e.KeyChar == '*')
            {
                ApplyBinaryOperator("×");
                e.Handled = true;
            }
            else if (e.KeyChar == '/')
            {
                ApplyBinaryOperator("÷");
                e.Handled = true;
            }
            else if (e.KeyChar == '%')
            {
                ApplyPercentage();
                e.Handled = true;
            }
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Alt && (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down))
            {
                ToggleAlwaysOnTop();
                e.Handled = true;
                return;
            }

            if (e.KeyCode == Keys.Enter)
            {
                OnCalculatorButtonClick("=");
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Back)
            {
                Backspace();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                if (menuOpen)
                    CloseNavMenu();
                else
                    OnCalculatorButtonClick("C");
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Delete)
            {
                OnCalculatorButtonClick("CE");
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.P)
            {
                ShowPrintReportOptions();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.H)
            {
                ToggleHistoryPanel();
                e.Handled = true;
            }
        }

        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
