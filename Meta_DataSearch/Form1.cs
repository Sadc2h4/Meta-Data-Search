using MetadataExtractor;
using MetadataExtractor.Formats.Exif;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Button = System.Windows.Forms.Button;
using Formatting = Newtonsoft.Json.Formatting;
using TextBox = System.Windows.Forms.TextBox;


namespace Meta_DataSearch
{
    public partial class Form1 : Form
    {
        // ====== 画面UI周り ======
        private Button btnOpen;
        private Button btnSaveRaw;        
        private CheckBox chkShowRaw;    // raw表示
        private CheckBox chkGridView;   // 表形式
        private TextBox txtOutput;
        private DataGridView grid;
        private Label lblHint;
        private const int RAW_BYTES_LIMIT = 4096; // rawでHEX表示する最大バイト数（大きすぎるICC/XMP対策）

        private PrivateFontCollection _fonts = new PrivateFontCollection();
        private Font _gridItemBoldFont;

        // ====== 直近結果の共有状態 ======
        private List<SdInfo> _lastInfos = new List<SdInfo>();
        private List<Dictionary<string, object>> _lastMetas = new List<Dictionary<string, object>>();

        // ===== 各パーツの色指定 =====
        static readonly Color BluePrimary = Color.FromArgb(0x3B, 0x82, 0xF6); // ボタン青
        static readonly Color OutputBg = ColorTranslator.FromHtml("#303136"); // txtOutput 背景
        static readonly Color WhiteText = Color.White;

        private readonly Dictionary<Button, Image> _buttonIconCache = new Dictionary<Button, Image>();

        // ===== 右上プレビュー =====
        private PictureBox _picPreview;
        private const int PreviewEdge = 64;
        private const int PreviewMin = 48;

        public Form1()
        {
            InitializeComponent();
            BuildUi();
            ApplyMinimalColors();
            ApplyGridColors();

            AllowDrop = true;
            DragEnter += Form1_DragEnter;
            DragDrop += Form1_DragDrop;

        }
        private void Form1_Load(object sender, EventArgs e) { }

        // =======================================================================================
        // UI 構築
        // =======================================================================================
        private void BuildUi()
        {
            Text = "Meta_DataSearch.exe";
            Width = 800;
            Height = 570;

            btnOpen = new Button { Text = "  Select image", Left = 12, Top = 12, Width = 160, Height = 36 };
            btnOpen.Click += BtnOpen_Click;

            btnSaveRaw = new Button { Text = "  Save raw data", Left = 180, Top = 12, Width = 180, Height = 36 };
            btnSaveRaw.Click += (s, e) => SaveRawJsonForCurrent();
            Controls.AddRange(new Control[] { btnOpen, btnSaveRaw, chkShowRaw, chkGridView, lblHint, txtOutput, grid });

            MakeRounded(btnOpen);
            MakeRounded(btnSaveRaw);

            Image imgSelect = Properties.Resources.Select_image_icon;
            Image imgSave = Properties.Resources.Raw_export_icon;

            StylePrimaryButton(btnOpen, imgSelect);
            StylePrimaryButton(btnSaveRaw, imgSave);

            chkShowRaw = new CheckBox { Text = "Display raw data", Left = 400, Top = 17, Width = 150, Checked = false };
            chkGridView = new CheckBox { Text = "Table mode", Left = 560, Top = 17, Width = 120, Checked = false };
            chkGridView.CheckedChanged += (_, __) => ToggleViewAndRender();
            chkShowRaw.CheckedChanged += (_, __) =>
            {
                if (_lastInfos.Count == 0) return;
                if (chkGridView.Checked) RenderGrid();  // グリッド表示中はそのまま再描画（rawはテキスト側で反映）
                else RenderText();                      // テキスト表示中ならRAWのON/OFFが即反映
                UpdatePreview(toolStripStatusLabel_imagePath.Text);
            };


            lblHint = new Label
            {
                Text = "PNG/JPG/WebP files can be loaded directly via drag-and-drop.",
                Left = 12,
                Top = 57,
                AutoSize = true
            };

            txtOutput = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                Left = 12,
                Top = 80,
                Width = ClientSize.Width - 24,
                Height = ClientSize.Height - 112,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Consolas", 10)
                //BackColor = "#303136"
            };

            grid = new DataGridView
            {
                Left = 12,
                Top = 80,
                Width = ClientSize.Width - 24,
                Height = ClientSize.Height - 112,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                DefaultCellStyle = { WrapMode = DataGridViewTriState.True }, // 折返す
                Visible = false
            };
            grid.CellContentClick += Grid_CellContentClick;
            grid.KeyDown += Grid_KeyDown;

            // 列構成：項目 / 値 / weight / コピー
            grid.Columns.Clear();
            var colItem = new DataGridViewTextBoxColumn
            {
                Name = "Item",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells,
                MinimumWidth = 120
            };
            var colValue = new DataGridViewTextBoxColumn
            {
                Name = "Contents",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, // 残り幅いっぱい
                DefaultCellStyle = { WrapMode = DataGridViewTriState.True } // 長文折返し
            };
            
            var colWeight = new DataGridViewTextBoxColumn
            {
                Name = "weight",
                HeaderText = "weight",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells,
                MinimumWidth = 60,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            };
            var colCopy = new DataGridViewButtonColumn
            {
                Name = "copy",
                Text = "copy",
                UseColumnTextForButtonValue = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            };
            grid.Columns.AddRange(colItem, colValue, colWeight, colCopy);
            grid.CellFormatting += Grid_CellFormatting;

            // ミニプレビューを作成
            _picPreview = new PictureBox
            {
                Width = PreviewEdge,
                Height = PreviewEdge,
                Top = 8,
                Left = this.ClientSize.Width - PreviewEdge - 12, // 右端から12px
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(0x30, 0x31, 0x36),     // ダーク面と同系
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.None
            };
            this.Controls.Add(_picPreview);
            // ボーダー風に見せる
            _picPreview.Paint += (s, e) =>
            {
                using (var pen = new Pen(Color.FromArgb(70, 255, 255, 255)))
                    e.Graphics.DrawRectangle(pen, 0, 0, _picPreview.Width - 1, _picPreview.Height - 1);
            };

            Controls.AddRange(new Control[] { btnOpen, btnSaveRaw, chkShowRaw, chkGridView, lblHint, txtOutput, grid });


        }

        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // ヘッダは除外、実セルのみ
            if (e.RowIndex < 0) return;

            // 「項目」列だけ太字にする
            var col = grid.Columns[e.ColumnIndex];
            if (col != null && col.Name == "Item")
            {
                // フォント作成は一度だけ
                if (_gridItemBoldFont == null ||
                    _gridItemBoldFont.FontFamily != grid.Font.FontFamily ||
                    Math.Abs(_gridItemBoldFont.Size - grid.Font.Size) > 0.1f)
                {
                    _gridItemBoldFont = new Font(grid.Font, FontStyle.Bold);
                }
                e.CellStyle.Font = _gridItemBoldFont;
            }
        }


        // すべての子コントロールを再帰列挙
        private static IEnumerable<Control> Walk(Control root)
        {
            foreach (Control c in root.Controls)
            {
                yield return c;
                foreach (var cc in Walk(c)) yield return cc;
            }
        }

        // ご指定どおりの“軽量テーマ”を適用（ラベル／ボタン／txtOutputのみ）
        // Label/CheckBox/Buttons/txtOutput の最小着色
        private void ApplyMinimalColors()
        {
            // Label → 白
            foreach (var lbl in Walk(this).OfType<Label>())
                lbl.ForeColor = WhiteText;

            // CheckBox → 白（BackColor はフォーム背景に合わせる）
            foreach (var cb in Walk(this).OfType<CheckBox>())
            {
                cb.ForeColor = WhiteText;
                cb.BackColor = this.BackColor;   // 透過で黒が見える場合は BackColor をフォーム色に
            }

            // Button → 青地＋白文字（UseVisualStyleBackColor をOFF）
            foreach (var btn in Walk(this).OfType<Button>())
            {
                btn.UseVisualStyleBackColor = false;
                btn.BackColor = BluePrimary;
                btn.ForeColor = WhiteText;
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 0;
            }

            // txtOutput → 背景#303136 / 文字白
            txtOutput.BackColor = OutputBg;
            txtOutput.ForeColor = WhiteText;
        }

        // DataGridView の背景/文字色を #303136 / 白 に
        private void ApplyGridColors()
        {
            grid.BackgroundColor = OutputBg;
            grid.BorderStyle = BorderStyle.FixedSingle;

            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersDefaultCellStyle.BackColor = OutputBg;
            grid.ColumnHeadersDefaultCellStyle.ForeColor = WhiteText;

            grid.DefaultCellStyle.BackColor = OutputBg;
            grid.DefaultCellStyle.ForeColor = WhiteText;
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(70, 90, 140);
            grid.DefaultCellStyle.SelectionForeColor = WhiteText;

            grid.AlternatingRowsDefaultCellStyle.BackColor = OutputBg; // 縞なし（軽量）
            grid.AlternatingRowsDefaultCellStyle.ForeColor = WhiteText;

            grid.RowHeadersDefaultCellStyle.BackColor = OutputBg;
            grid.RowHeadersDefaultCellStyle.ForeColor = WhiteText;
            grid.GridColor = Color.FromArgb(70, 70, 76);
        }

        // =======================================================================================
        // 表示モード切替時の処理
        // =======================================================================================
        private void ToggleViewAndRender()
        {
            var toGrid = chkGridView.Checked;
            grid.Visible = toGrid;
            txtOutput.Visible = !toGrid;

            if (_lastInfos.Count > 0)
            {
                if (toGrid) { RenderGrid(); ApplyGridColors(); } 
                else { RenderText(); }
                UpdatePreview(toolStripStatusLabel_imagePath.Text);
            }
        }

        // =======================================================================================
        // 角丸リージョンを作成してボタンの角を丸くする
        // =======================================================================================
        private static System.Drawing.Drawing2D.GraphicsPath RoundRect(Rectangle r, int radius)
        {
            int d = radius * 2;
            var gp = new System.Drawing.Drawing2D.GraphicsPath();
            gp.AddArc(r.X, r.Y, d, d, 180, 90);
            gp.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            gp.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            gp.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            gp.CloseFigure();
            return gp;
        }
        private void MakeRounded(Button b, int radius = 10)
        {
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 0;
            b.UseVisualStyleBackColor = false;
            b.Region = new Region(RoundRect(b.ClientRectangle, radius));
            b.Resize += (s, e) => b.Region = new Region(RoundRect(b.ClientRectangle, radius));
        }

        // =======================================================================================
        //　ボタンのスタイルを設定する
        // =======================================================================================
        private void StylePrimaryButton(Button b, Image icon = null)
        {            
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 0;
            b.BackColor = BluePrimary;
            b.ForeColor = WhiteText;
            b.Font = new Font(b.Font, FontStyle.Bold); // 起動直後から太字
            b.Padding = new Padding(12, 6, 12, 6);
            b.Height = 36; // 念のため固定
            if (icon != null)
            {                
                b.Resize += (s, e) => ApplyButtonIcon(b, icon); // フォームresize時に追従するようにしたいので
                ApplyButtonIcon(b, icon);

                b.ImageAlign = ContentAlignment.MiddleLeft;
                b.TextImageRelation = TextImageRelation.ImageBeforeText;
                b.Padding = new Padding(10, 6, 12, 6);
            }
            MakeRounded(b);
        }

        private static Image ResizeImageHighQuality(Image src, int w, int h, int vShift = 0)
        {
            var bmp = new Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                var scale = Math.Min((float)w / src.Width, (float)h / src.Height);
                int dw = Math.Max(1, (int)Math.Round(src.Width * scale));
                int dh = Math.Max(1, (int)Math.Round(src.Height * scale));
                int dx = (w - dw) / 2;
                int dy = (h - dh) / 2 + vShift - 3;        // アイコンを上下にずらす（上に行くほどマイナス）

                // はみ出し防止の軽いクランプ（±4px 目安）
                int clamp = 4;
                dy = Math.Max(-(clamp), Math.Min(clamp, dy));

                g.DrawImage(src, new Rectangle(dx, dy, dw, dh));
            }
            return bmp;
        }
        // ボタン高さに応じてアイコンを作り直してセット（左に表示）
        private void ApplyButtonIcon(Button b, Image src)
        {
            if (src == null) { b.Image = null; return; }

            // 上下の余白。アイコンが切れる場合は 9〜10 に
            var topBottom = 7;
            b.Padding = new Padding(12, topBottom, 14, topBottom);

            // 実際に使える高さからアイコンサイズを算出
            int availH = Math.Max(16, b.ClientSize.Height - b.Padding.Vertical);

            // 文字ベースラインと揃うように、ほんの少しだけ上へ寄せる
            int vShift = -(int)Math.Round(b.DeviceDpi / 144.0);  // 96dpi:-1 / 144dpi:-1 / 192dpi:-2 ぐらい

            Image old;
            if (_buttonIconCache.TryGetValue(b, out old) && old != null) old.Dispose();
            var resized = ResizeImageHighQuality(src, availH, availH, vShift);
            _buttonIconCache[b] = resized;

            b.Image = resized;
            b.ImageAlign = ContentAlignment.MiddleLeft;
            b.TextImageRelation = TextImageRelation.ImageBeforeText;
            b.TextAlign = ContentAlignment.MiddleLeft;
        }

        // =======================================================================================
        // ファイル選択ボタン押下時の処理
        // =======================================================================================
        private void BtnOpen_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog
            {
                Title = "画像を選択",
                Filter = "Images|*.png;*.jpg;*.jpeg;*.webp|All files|*.*",
                Multiselect = true
            })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                // ダイアログの RCW を解放させてから処理したいので一旦キャプチャ
                var files = dlg.FileNames;
                var first = dlg.FileName;

                // UI メッセージループに戻した後に実行（COM の再入を回避）
                BeginInvoke(new Action(async () =>
                {
                    SetBusy(true, "Reading metadata...");
                    try
                    {
                        toolStripStatusLabel_imagePath.Text = first;
                        ProcessFiles(files);
                        UpdatePreview(first);
                        await PostProcessMissingMetaAsync();   // OK
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "AI生成でエラー: " + ex.Message);
                    }
                    finally
                    {
                        SetBusy(false, "Ready");
                    }
                }));
            }
        }
        // =======================================================================================
        // ドラッグ&ドロップ処理
        // =======================================================================================
        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }
        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files == null || files.Length == 0) return;

            var first = files[0];

            // Drag&Drop は内部で COM の同期呼び出し中なので、必ず一拍置く
            BeginInvoke(new Action(async () =>
            {
                SetBusy(true, "Reading metadata...");
                try
                {
                    toolStripStatusLabel_imagePath.Text = first;
                    ProcessFiles(files);
                    UpdatePreview(first);
                    await PostProcessMissingMetaAsync();   // OK
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "AI生成でエラー: " + ex.Message);
                }
                finally
                {
                    SetBusy(false, "Ready");
                }
            }));
        }

        // =======================================================================================
        // 実行本体（結果を共有状態に格納し、両ビューへ描画）
        // =======================================================================================
        private void ProcessFiles(IEnumerable<string> paths)
        {
            _lastInfos.Clear();
            _lastMetas.Clear();

            foreach (var p in paths)
            {
                try
                {
                    var meta = Extract(p);
                    var info = BuildSdInfo(meta, p);
                    _lastMetas.Add(meta);
                    _lastInfos.Add(info);
                }
                catch (Exception ex)
                {
                    // 失敗ファイルも表示に出す
                    _lastInfos.Add(new SdInfo { FileName = $"{Path.GetFileName(p)} [error: {ex.Message}]" });
                    _lastMetas.Add(new Dictionary<string, object> { { "error", ex.Message } });
                }
            }

            if (chkGridView.Checked) RenderGrid();
            else RenderText();　// テキストボックスに内容を反映
            UpdatePreview(toolStripStatusLabel_imagePath.Text);

        }


        // =======================================================================================
        // 表形式描画 / テキスト描画 / モデル等をリンク化
        // =======================================================================================
        private void RenderGrid()
        {
            grid.SuspendLayout();
            grid.Rows.Clear();

            foreach (var x in _lastInfos)
            {
                foreach (var row in BuildGridRows(x))
                {
                    int idx = grid.Rows.Add(row.Item, row.Value, row.Weight, "copy");
                    if (!string.IsNullOrEmpty(row.Url))
                    {
                        var link = new DataGridViewLinkCell();
                        link.Value = row.Value;
                        link.Tag = row.Url;
                        link.LinkColor = Color.FromArgb(120, 215, 255);       // 見やすい水色
                        link.VisitedLinkColor = link.LinkColor;
                        link.ActiveLinkColor = Color.White;
                        link.TrackVisitedState = false;
                        grid.Rows[idx].Cells["Contents"] = link;
                    }
                }
                grid.Rows.Add("――――――", "――――――", "", "");
            }
            grid.ResumeLayout();
            grid.AutoResizeRows();

            ApplyGridColors();       // 行追加後に色を再適用
        }

        // =======================================================================================
        // テキストボックスに文字列を反映・描画
        // =======================================================================================
        private void RenderText()
        {            
            txtOutput.Clear();
            for (int i = 0; i < _lastInfos.Count; i++)
            {
                var showRaw = chkShowRaw.Checked ? _lastMetas[i] : null;
                if (chkShowRaw.Checked)
                {
                    var sb = new StringBuilder();
                    foreach (var x in _lastInfos)
                    {
                        if (string.IsNullOrEmpty(x.FullPath) || !File.Exists(x.FullPath))
                            continue;

                        sb.AppendLine(DumpAllMetadataToJson(x.FullPath));
                        sb.AppendLine();
                        sb.AppendLine("----");
                        sb.AppendLine();
                    }
                    txtOutput.Text = sb.ToString();
                    return;
                }

                txtOutput.AppendText(FormatReport(_lastInfos[i], showRaw));
                txtOutput.AppendText(Environment.NewLine + Environment.NewLine);
            }
        }

        // =======================================================================================
        // コピー列（ボタン）クリック → 値列をコピー
        // =======================================================================================
        private void Grid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var colName = grid.Columns[e.ColumnIndex].Name;

            // copy ボタン → 「Contents」列の文字列をコピー
            if (colName == "copy")
            {
                var val = Convert.ToString(grid.Rows[e.RowIndex].Cells["Contents"].Value ?? "");
                if (!string.IsNullOrEmpty(val)) Clipboard.SetText(val);
                return;
            }

            // Contents 列がリンクなら URL を開く
            if (colName == "Contents")
            {
                var cell = grid.Rows[e.RowIndex].Cells[e.ColumnIndex] as DataGridViewLinkCell;
                var url = cell?.Tag as string;
                if (!string.IsNullOrEmpty(url))
                {
                    try
                    {
                        System.Diagnostics.Process.Start(
                            new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true });
                    }
                    catch { /* 無視 */ }
                }
            }
        }


        // =======================================================================================
        // Ctrl+C で選択セルの内容をコピー
        // =======================================================================================
        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                var cell = grid.CurrentCell;
                var txt = cell?.Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(txt)) Clipboard.SetText(txt);
                e.Handled = true;
            }
        }

        // =======================================================================================
        // 構造化モデル
        // =======================================================================================
        // 既存を置き換え
        private class LoraInfo 
        { 
            public string Name; 
            public string Weight; 
            public string Hash;
            public string VersionId;   // modelVersionId を保持
            public string ModelId;
        }  
        
        private class SdInfo
        {
            public string FileName;
            public string Checkpoint;
            public List<LoraInfo> Loras = new List<LoraInfo>();
            public List<EmbeddingInfo> Embeddings = new List<EmbeddingInfo>();
            public string Prompt;
            public string NegativePrompt;
            public string Seed;
            public string Steps;
            public string Cfg;
            public string Denoise;
            public string Sampler;
            public int? Width;
            public int? Height;
            public string Scale;
            public string ModelHash;   // リンク製作用のhashを追加
            public string SizeText => (Width.HasValue && Height.HasValue) ? $"{Width}x{Height}" : "";
            public string FullPath;
            public string CheckpointVersionId;
            public string CheckpointModelId;
        }

        private class EmbeddingInfo
        {
            public string Name;
            public string Weight; // 任意（<embedding:name:0.8> の 0.8 部分）
            public string Hash;   // TI hashes などから付与
            public string VersionId;
            public string ModelId;
        }

        private class GridRow
        {
            public string Item;
            public string Value;
            public string Weight;
            public string Url;
        }
        private GridRow R(string a, string b, string w = "", string url = null)
                        => new GridRow { Item = a, Value = b, Weight = w, Url = url };

        // meta(dict) + 画像ファイル から SdInfo を構築（Comfy優先）
        private SdInfo BuildSdInfo(Dictionary<string, object> meta, string imagePath)
        {
            var info = new SdInfo
            {
                FileName = Path.GetFileName(imagePath),
                FullPath = imagePath
            };

            // 画像実寸
            try { using (var img = Image.FromFile(imagePath)) { info.Width = img.Width; info.Height = img.Height; } } catch { }

            // まず A1111 の値を控えとして取得（直接は採用しない）
            string aPrompt = null, aNeg = null, aSteps = null, aCfg = null, aSampler = null, aDenoise = null, aScale = null;

            var parsed = meta.ContainsKey("parsed") ? meta["parsed"] as Dictionary<string, string> : null;
            if (parsed != null)
            {
                parsed.TryGetValue("negative_prompt", out aNeg);

                string ptmp;
                if (parsed.TryGetValue("prompt", out ptmp))
                {
                    if (!IsLikelyComfyJson(ptmp)) aPrompt = ptmp; // Comfy JSON っぽいなら廃棄
                }
                parsed.TryGetValue("sampler", out aSampler);
                parsed.TryGetValue("steps", out aSteps);
                parsed.TryGetValue("cfg_scale", out aCfg);
                parsed.TryGetValue("denoising_strength", out aDenoise);
                parsed.TryGetValue("hires_upscale", out aScale);

                // size は実寸へ反映
                string aSize;
                if (parsed.TryGetValue("size", out aSize))
                {
                    var m = Regex.Match(aSize, @"(\d+)\s*[xX]\s*(\d+)");
                    if (m.Success) { info.Width = int.Parse(m.Groups[1].Value); info.Height = int.Parse(m.Groups[2].Value); }
                }
                // A1111 の checkpoint は使わない（Comfy からのみセット）
            }

            // “Prompt 全量 JSON” を集める（Comfy を最優先）
            var soup = new StringBuilder();
            if (parsed != null)
            {
                string maybeComfy;
                if (parsed.TryGetValue("prompt", out maybeComfy) && IsLikelyComfyJson(maybeComfy))
                    soup.AppendLine(maybeComfy);
            }
            string a1111Params = null;
            object rawObj;
            if(meta.TryGetValue("raw", out rawObj) && rawObj is Dictionary<string, object>)
{
                var raw = (Dictionary<string, object>)rawObj;

                object xmp; if (raw.TryGetValue("xmp", out xmp) && xmp is string) soup.AppendLine((string)xmp);

                object png;
                if (raw.TryGetValue("png", out png) && png is Dictionary<string, string>)
                {
                    var pngText = (Dictionary<string, string>)png;

                    // 既存：全値を Comfy 推定用の soup に入れておく
                    foreach (var kv in pngText) soup.AppendLine(kv.Value);

                    // A1111 "parameters" を拾う（キーの大小/別名をケア）
                    foreach (var key in new[] { "parameters", "Parameters", "Description", "Comment" })
                    {
                        string tmp;
                        if (pngText.TryGetValue(key, out tmp) && !string.IsNullOrWhiteSpace(tmp))
                        {
                            a1111Params = tmp;
                            break;
                        }
                    }
                }

                object exif; if (raw.TryGetValue("exif", out exif) && exif is Dictionary<string, string>)
                    foreach (var kv in (Dictionary<string, string>)exif) soup.AppendLine(kv.Value);
            }
            // PNG に見つからない場合は EXIF / XMP 側からも拾って補完
            if (string.IsNullOrWhiteSpace(a1111Params) && meta.TryGetValue("raw", out rawObj) && rawObj is Dictionary<string, object>)
            {
                var raw = (Dictionary<string, object>)rawObj;

                var sbA = new StringBuilder();
                object ex; if (raw.TryGetValue("exif", out ex) && ex is Dictionary<string, string>)
                    foreach (var kv in (Dictionary<string, string>)ex) sbA.AppendLine(kv.Value);

                object xmp; if (raw.TryGetValue("xmp", out xmp) && xmp is string)
                    sbA.AppendLine((string)xmp);

                var candidate = sbA.ToString();
                if (HasSdMarkers(candidate))
                    a1111Params = candidate;
            }

            // Comfy JSON から厳密に抽出
            ExtractFromComfyJson(soup.ToString(), info);
            ExtractFromA1111Parameters(a1111Params, info);
            ExtractCivitaiResourcesFromText(a1111Params, info);

            object rawObj3;
            if (meta.TryGetValue("raw", out rawObj3) && rawObj3 is Dictionary<string, object>)
            {
                var rawDictTmp = (Dictionary<string, object>)rawObj3;
                object exifObj;
                if (rawDictTmp.TryGetValue("exif", out exifObj) && (exifObj is Dictionary<string, string>))
                {
                    foreach (var kv in (Dictionary<string, string>)exifObj)
                    {
                        var v = kv.Value;
                        if (!string.IsNullOrEmpty(v) &&
                            v.IndexOf("Civitai resources:", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            ExtractCivitaiResourcesFromText(v, info);
                        }
                    }
                }
            }

            // 未充足分は A1111 で補完する
            if (string.IsNullOrEmpty(info.Prompt) && !string.IsNullOrEmpty(aPrompt)) info.Prompt = aPrompt;
            if (string.IsNullOrEmpty(info.NegativePrompt) && !string.IsNullOrEmpty(aNeg)) info.NegativePrompt = aNeg;
            if (string.IsNullOrEmpty(info.Steps) && !string.IsNullOrEmpty(aSteps)) info.Steps = aSteps;
            if (string.IsNullOrEmpty(info.Cfg) && !string.IsNullOrEmpty(aCfg)) info.Cfg = aCfg;
            if (string.IsNullOrEmpty(info.Sampler) && !string.IsNullOrEmpty(aSampler)) info.Sampler = aSampler;
            if (string.IsNullOrEmpty(info.Denoise) && !string.IsNullOrEmpty(aDenoise)) info.Denoise = aDenoise;
            if (string.IsNullOrEmpty(info.Scale) && !string.IsNullOrEmpty(aScale)) info.Scale = aScale;
            
            
            // LoRA（プロンプト/ネガ両方から）
            if (info.Loras.Count == 0)
            {
                var basePrompt = (info.Prompt ?? aPrompt) ?? "";
                var baseNeg = (info.NegativePrompt ?? aNeg) ?? "";
                foreach (var l in ExtractLorasFromPrompt(basePrompt + "\n" + baseNeg))
                    AddLoraUnique(info, l.Name, l.Weight);
            }

            // Checkpoint（Model名/Hashから）
            if (string.IsNullOrEmpty(info.Checkpoint) && parsed != null)
            {
                string modelName = null, modelHash = null;
                parsed.TryGetValue("model", out modelName);
                parsed.TryGetValue("model_hash", out modelHash);
                if (!string.IsNullOrEmpty(modelName))
                    info.Checkpoint = string.IsNullOrEmpty(modelHash) ? modelName : (modelName + " (" + modelHash + ")");
            }

            // Model hash（A1111 由来）も保持
            if (string.IsNullOrEmpty(info.ModelHash) && parsed != null)
            {
                string mh;
                if (parsed.TryGetValue("model_hash", out mh) && !string.IsNullOrWhiteSpace(mh))
                    info.ModelHash = mh;
            }

            // LoRAハッシュ（raw→png→Steps の "Lora hashes:" から付与）
            Dictionary<string, object> rawDict = null;   // ← 名前を raw ではなく rawDict に
            object rawObj2;
            if (meta.TryGetValue("raw", out rawObj2) && (rawObj2 is Dictionary<string, object>))
                rawDict = (Dictionary<string, object>)rawObj2;

            if (rawDict != null)
            {
                object pngObj;
                if (rawDict.TryGetValue("png", out pngObj) && (pngObj is Dictionary<string, string>))
                {
                    var png = (Dictionary<string, string>)pngObj;
                    string stepsLine;
                    if (png.TryGetValue("Steps", out stepsLine))
                    {
                        var map = ParseLoraHashesFromSteps(stepsLine); // 例: { name: hash, ... }
                        if (map.Count > 0)
                        {
                            // LoRAがまだ無ければ、ハッシュの名前だけでも作成
                            if (info.Loras.Count == 0)
                            {
                                foreach (var kv in map)
                                    AddLoraUnique(info, kv.Key, null);   // Weightは不明なので null
                            }
                            // 既存LoRAにハッシュを紐付け
                            foreach (var l in info.Loras)
                            {
                                string h;
                                if (l.Hash == null && map.TryGetValue(l.Name, out h))
                                    l.Hash = h;
                            }
                        }
                    }
                }
            }

            // raw["png"]["Steps"] から TI hashes を拾う
            if (rawDict != null && rawDict.TryGetValue("png", out object pngObj2) && (pngObj2 is Dictionary<string, string>))
            {
                var png2 = (Dictionary<string, string>)pngObj2;
                string stepsLine2;
                if (png2.TryGetValue("Steps", out stepsLine2))
                {
                    var tiMap = ParseTiHashesFromSteps(stepsLine2);
                    if (tiMap.Count > 0)
                    {
                        if (info.Embeddings.Count == 0)
                            foreach (var kv in tiMap) AddEmbeddingUnique(info, kv.Key, null);

                        foreach (var e in info.Embeddings)
                        {
                            string h;
                            if (e.Hash == null && tiMap.TryGetValue(e.Name, out h))
                                e.Hash = h;
                        }
                    }
                }
            }


            return info;
        }

        // =======================================================================================
        // A1111の"prompt"にComfyのJSONが入っているか確認する処理
        // =======================================================================================
        private bool IsLikelyComfyJson(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            if (s.IndexOf("\\u0022class_type\\u0022", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (s.IndexOf("class_type", StringComparison.OrdinalIgnoreCase) >= 0 &&
                s.IndexOf("CheckpointLoaderSimple", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }

        // =======================================================================================
        // ---------- ComfyUI JSON 解析 ----------
        // =======================================================================================
        private void ExtractFromComfyJson(string s, SdInfo info)
        {
            if (string.IsNullOrWhiteSpace(s)) return;
            foreach (var obj in FindJsonObjects(s))
                ParseComfyGraph(obj, info);
        }
        private IEnumerable<JObject> FindJsonObjects(string text)
        {
            var list = new List<JObject>();
            int n = text.Length;
            for (int i = 0; i < n; i++)
            {
                if (text[i] != '{') continue;
                bool inStr = false; bool esc = false; int depth = 0;
                for (int j = i; j < n; j++)
                {
                    char c = text[j];
                    if (inStr)
                    {
                        if (esc) esc = false;
                        else if (c == '\\') esc = true;
                        else if (c == '"') inStr = false;
                    }
                    else
                    {
                        if (c == '"') inStr = true;
                        else if (c == '{') depth++;
                        else if (c == '}')
                        {
                            depth--;
                            if (depth == 0)
                            {
                                var candidate = text.Substring(i, j - i + 1);
                                try
                                {
                                    var jo = JsonConvert.DeserializeObject<JObject>(candidate);
                                    if (jo != null) list.Add(jo);
                                }
                                catch { }
                                i = j;
                                break;
                            }
                        }
                    }
                }
            }
            return list;
        }

        // "Civitai resources: [ ... ]" のようなラベルの直後にある JSON 配列を抜き出す
        private static string FindJsonArrayAfterLabel(string text, string label)
        {
            if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(label)) return null;
            int p = text.IndexOf(label, StringComparison.OrdinalIgnoreCase);
            if (p < 0) return null;

            // ラベルの後ろから '[' を探す
            int i = text.IndexOf('[', p);
            if (i < 0) return null;

            // 角括弧の対応を取りながら配列全体を切り出す
            int depth = 0;
            for (int j = i; j < text.Length; j++)
            {
                char c = text[j];
                if (c == '[') depth++;
                else if (c == ']')
                {
                    depth--;
                    if (depth == 0)
                    {
                        return text.Substring(i, j - i + 1);
                    }
                }
            }
            return null; // 閉じカッコ見つからず
        }

        // Civitai resources の JSON を解析して SdInfo に反映
        private void ExtractCivitaiResourcesFromText(string text, SdInfo info)
        {
            if (string.IsNullOrWhiteSpace(text) || info == null) return;

            var json = FindJsonArrayAfterLabel(text, "Civitai resources:");
            if (string.IsNullOrEmpty(json)) return;

            try
            {
                var arr = Newtonsoft.Json.Linq.JArray.Parse(json);
                foreach (var tok in arr)
                {
                    var o = tok as Newtonsoft.Json.Linq.JObject;
                    if (o == null) continue;

                    var type = (string)(o["type"] ?? "");
                    var modelName = (string)(o["modelName"] ?? o["name"] ?? "");
                    var verName = (string)(o["modelVersionName"] ?? "");
                    var verId = (o["modelVersionId"] ?? o["versionId"])?.ToString(); // 追加
                    var modelId = (o["modelId"])?.ToString();                           // 追加
                    var weight = o["weight"] != null ? o["weight"].ToString() : null;

                    if (type.Equals("checkpoint", StringComparison.OrdinalIgnoreCase))
                    {
                        if (string.IsNullOrEmpty(info.Checkpoint))
                            info.Checkpoint = string.IsNullOrEmpty(verName) ? modelName : $"{modelName} ({verName})";
                        if (!string.IsNullOrEmpty(verId)) info.CheckpointVersionId = verId;  // 追加
                        if (!string.IsNullOrEmpty(modelId)) info.CheckpointModelId = modelId;// 追加
                    }
                    else if (type.Equals("lora", StringComparison.OrdinalIgnoreCase) ||
                             type.Equals("lycoris", StringComparison.OrdinalIgnoreCase))
                    {
                        AddLoraUnique(info, modelName, weight);
                        var last = info.Loras.LastOrDefault();
                        if (last != null) { last.VersionId = verId; last.ModelId = modelId; } // 追加
                    }
                    else if (type.Equals("embed", StringComparison.OrdinalIgnoreCase) ||
                             type.Equals("embedding", StringComparison.OrdinalIgnoreCase))
                    {
                        AddEmbeddingUnique(info, modelName, weight);
                        var last = info.Embeddings.LastOrDefault();
                        if (last != null) { last.VersionId = verId; last.ModelId = modelId; } // 追加
                    }

                }
            }
            catch
            {
                // JSON パースに失敗した時は黙ってスキップ（他経路で拾える場合が多い）
            }
        }


        private void ParseComfyGraph(JObject root, SdInfo info)
        {
            // extraMetadata（二重デコード）
            var extraMetaToken = root["extraMetadata"];
            if (extraMetaToken != null && extraMetaToken.Type == JTokenType.String)
            {
                try
                {
                    var inner = (string)extraMetaToken; // {"\u0022prompt\u0022": ...}
                    var innerObj = JsonConvert.DeserializeObject<JObject>(inner);
                    if (innerObj != null)
                    {
                        if (string.IsNullOrEmpty(info.Prompt)) info.Prompt = CleanText((string)innerObj["prompt"]);
                        if (string.IsNullOrEmpty(info.NegativePrompt)) info.NegativePrompt = CleanText((string)innerObj["negativePrompt"]);
                        if (string.IsNullOrEmpty(info.Steps) && innerObj["steps"] != null) info.Steps = innerObj["steps"].ToString();
                        if (string.IsNullOrEmpty(info.Cfg) && innerObj["cfgScale"] != null) info.Cfg = innerObj["cfgScale"].ToString();
                        if (string.IsNullOrEmpty(info.Sampler) && innerObj["sampler"] != null) info.Sampler = innerObj["sampler"].ToString();
                    }
                }
                catch { }
            }

            foreach (var prop in root.Properties())
            {
                var node = prop.Value as JObject;   // C#7.3
                if (node == null) continue;

                var cls = node["class_type"] != null ? node["class_type"].ToString() : "";
                var inputs = node["inputs"] as JObject;

                // ---- Checkpoint（包含一致 & 複数キーに対応）----
                if (cls.IndexOf("checkpoint", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    var ckTok = inputs != null
                        ? (inputs["ckpt_name"] ?? inputs["ckpt"] ?? inputs["model"] ?? inputs["model_name"])
                        : null;
                    var ck = ckTok != null ? ckTok.ToString() : null;
                    if (!string.IsNullOrEmpty(ck) && string.IsNullOrEmpty(info.Checkpoint))
                        info.Checkpoint = TrimUrn(ck);
                }

                // ---- LoRA（名称ゆれ対応 + 配列 "loras" に対応）----
                if (cls.IndexOf("lora", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    cls.IndexOf("lyco", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (inputs != null)
                    {
                        // 単発型
                        var nameTok = inputs["lora_name"] ?? inputs["model"] ?? inputs["lora"];
                        var wTok = inputs["strength_model"] ?? inputs["strength"] ?? inputs["weight"];
                        if (nameTok != null)
                            AddLoraUnique(info, TrimUrn(nameTok.ToString()), wTok?.ToString());

                        // 配列型（LoraLoaderStacked等）
                        var arr = inputs["loras"] as JArray;
                        if (arr != null)
                        {
                            foreach (var it in arr.OfType<JObject>())
                            {
                                var n = it["lora_name"] ?? it["name"];
                                var w = it["strength_model"] ?? it["strength"] ?? it["weight"];
                                if (n != null) AddLoraUnique(info, TrimUrn(n.ToString()), w?.ToString());
                            }
                        }
                    }
                }

                // Sampler / Steps / CFG / Seed / Denoise
                if (cls.StartsWith("KSampler", StringComparison.OrdinalIgnoreCase))
                {
                    if (info.Seed == null && inputs != null && inputs["seed"] != null) info.Seed = inputs["seed"].ToString();
                    if (info.Steps == null && inputs != null && inputs["steps"] != null) info.Steps = inputs["steps"].ToString();
                    if (info.Cfg == null && inputs != null && inputs["cfg"] != null) info.Cfg = inputs["cfg"].ToString();
                    if (info.Sampler == null && inputs != null && inputs["sampler_name"] != null) info.Sampler = inputs["sampler_name"].ToString();
                    if (info.Denoise == null && inputs != null && inputs["denoise"] != null) info.Denoise = inputs["denoise"].ToString();
                }

                // Positive / Negative
                if (cls.IndexOf("CLIPTextEncode", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    var text = inputs != null ? inputs["text"]?.ToString() : null;
                    var title = node["_meta"] != null ? node["_meta"]["title"]?.ToString() : null;
                    if (!string.IsNullOrEmpty(text))
                    {
                        if (string.Equals(title, "Positive", StringComparison.OrdinalIgnoreCase) && info.Prompt == null)
                            info.Prompt = CleanText(text);
                        else if (string.Equals(title, "Negative", StringComparison.OrdinalIgnoreCase) && info.NegativePrompt == null)
                            info.NegativePrompt = CleanText(text);
                    }
                }

                // サイズ / スケール
                if (inputs != null)
                {
                    int v;
                    var wTok = inputs["width"]; var hTok = inputs["height"];
                    if (wTok != null && int.TryParse(wTok.ToString(), out v))
                        if (!info.Width.HasValue || v > info.Width) info.Width = v;
                    if (hTok != null && int.TryParse(hTok.ToString(), out v))
                        if (!info.Height.HasValue || v > info.Height) info.Height = v;

                    var scaleTok = inputs["scale"] ?? inputs["upscale"];
                    if (scaleTok != null && string.IsNullOrEmpty(info.Scale))
                        info.Scale = scaleTok.ToString();
                }
            }

            // extra.airs（補助）
            var airs = root["extra"] != null ? root["extra"]["airs"] as JArray : null;
            if (airs != null)
            {
                foreach (var t in airs)
                {
                    var s = t.ToString();
                    if (s.Contains(":checkpoint:") && string.IsNullOrEmpty(info.Checkpoint))
                        info.Checkpoint = TrimUrn(s);
                    else if (s.Contains(":lora:"))
                        AddLoraUnique(info, TrimUrn(s), null);
                }
            }
        }

        // =======================================================================================
        // LoRA 重複追加防止処理
        // =======================================================================================
        private void AddLoraUnique(SdInfo info, string name, string weight)
        {
            if (string.IsNullOrEmpty(name)) return;
            if (!info.Loras.Any(l => string.Equals(l.Name, name, StringComparison.OrdinalIgnoreCase)))
                info.Loras.Add(new LoraInfo { Name = name, Weight = weight });
        }

        // =======================================================================================
        // Embed 重複追加防止処理
        // =======================================================================================
        private void AddEmbeddingUnique(SdInfo info, string name, string weight)
        {
            if (string.IsNullOrWhiteSpace(name)) return;
            for (int i = 0; i < info.Embeddings.Count; i++)
            {
                if (string.Equals(info.Embeddings[i].Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrEmpty(info.Embeddings[i].Weight) && !string.IsNullOrEmpty(weight))
                        info.Embeddings[i].Weight = weight;
                    return;
                }
            }
            info.Embeddings.Add(new EmbeddingInfo { Name = name, Weight = weight });
        }
        private string TrimUrn(string urn)
        {
            if (string.IsNullOrEmpty(urn)) return urn;
            int idx = urn.LastIndexOf("civitai:", StringComparison.OrdinalIgnoreCase);
            if (idx >= 0) return urn.Substring(idx);
            idx = urn.LastIndexOf(':');
            return (idx >= 0 && idx + 1 < urn.Length) ? urn.Substring(idx + 1) : urn;
        }

        private string CleanText(string t)
        {
            if (string.IsNullOrEmpty(t)) return t;
            t = t.Replace("\\n", "\n").Replace("\\r", "\r");
            return t.Trim();
        }

        // 整形出力（テキスト）
        private string FormatReport(SdInfo x, Dictionary<string, object> showRaw)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"File : {x.FileName}");
            if (!string.IsNullOrEmpty(x.Checkpoint)) sb.AppendLine($"Checkpoint: {x.Checkpoint}");
            if (x.Loras.Count > 0)
                foreach (var l in x.Loras) sb.AppendLine($"Lora: {l.Name}   strength_model:{l.Weight}");
            sb.AppendLine();

            sb.AppendLine(new string('-', 100));
            sb.AppendLine("Prompt:");
            if (!string.IsNullOrWhiteSpace(x.Prompt)) sb.AppendLine(x.Prompt);
            sb.AppendLine(new string('-', 100));
            sb.AppendLine("Negative Prompt:");
            if (!string.IsNullOrWhiteSpace(x.NegativePrompt)) sb.AppendLine(x.NegativePrompt);
            sb.AppendLine(new string('-', 100));
            sb.AppendLine();

            if (!string.IsNullOrEmpty(x.Seed)) sb.AppendLine($"Seed:{x.Seed}");
            if (!string.IsNullOrEmpty(x.Steps)) sb.AppendLine($"Steps:{x.Steps}");
            if (!string.IsNullOrEmpty(x.Cfg)) sb.AppendLine($"cfg:{x.Cfg}");
            if (!string.IsNullOrEmpty(x.Denoise)) sb.AppendLine($"denoise:{x.Denoise}");
            if (!string.IsNullOrEmpty(x.Sampler)) sb.AppendLine($"Sampler:{x.Sampler}");
            if (!string.IsNullOrEmpty(x.SizeText)) sb.AppendLine($"Size:{x.SizeText}");
            if (!string.IsNullOrEmpty(x.Scale)) sb.AppendLine($"Scale:{x.Scale}");

            if (showRaw != null)
            {
                sb.AppendLine();
                sb.AppendLine(new string('=', 100));
                sb.AppendLine("RAW:");
                sb.AppendLine(JsonConvert.SerializeObject(showRaw, Formatting.Indented));
            }
            return sb.ToString();
        }

        // 表形式の行データ
        private List<GridRow> BuildGridRows(SdInfo x)
        {
            var rows = new List<GridRow>();
            rows.Add(R("File", x.FileName));
            // Checkpoint
            if (!string.IsNullOrEmpty(x.Checkpoint))
            {
                string urlCkpt;
                TryMakeCivitaiUrl(x.Checkpoint, x.ModelHash, x.CheckpointVersionId, out urlCkpt);
                rows.Add(R("Checkpoint", x.Checkpoint, "", urlCkpt));
            }

            // Lora
            if (x.Loras.Count > 0)
            {
                string url;
                TryMakeCivitaiUrl(x.Loras[0].Name, x.Loras[0].Hash, x.Loras[0].VersionId, out url);
                rows.Add(R("Lora", x.Loras[0].Name, x.Loras[0].Weight, url));
                for (int i = 1; i < x.Loras.Count; i++)
                {
                    TryMakeCivitaiUrl(x.Loras[i].Name, x.Loras[i].Hash, x.Loras[i].VersionId, out url);
                    rows.Add(R("", x.Loras[i].Name, x.Loras[i].Weight, url));
                }
            }

            // Embedding
            if (x.Embeddings.Count > 0)
            {
                string urlE;
                TryMakeCivitaiUrl(x.Embeddings[0].Name, x.Embeddings[0].Hash, x.Embeddings[0].VersionId, out urlE);
                rows.Add(R("Embedding", x.Embeddings[0].Name, x.Embeddings[0].Weight, urlE));
                for (int i = 1; i < x.Embeddings.Count; i++)
                {
                    TryMakeCivitaiUrl(x.Embeddings[i].Name, x.Embeddings[i].Hash, x.Embeddings[i].VersionId, out urlE);
                    rows.Add(R("", x.Embeddings[i].Name, x.Embeddings[i].Weight, urlE));
                }
            }


            rows.Add(R("Prompt", x.Prompt));
            rows.Add(R("Negative Prompt", x.NegativePrompt));

            if (!string.IsNullOrEmpty(x.Seed)) rows.Add(R("Seed", x.Seed));
            if (!string.IsNullOrEmpty(x.Steps)) rows.Add(R("Steps", x.Steps));
            if (!string.IsNullOrEmpty(x.Cfg)) rows.Add(R("cfg", x.Cfg));
            if (!string.IsNullOrEmpty(x.Denoise)) rows.Add(R("denoise", x.Denoise));
            if (!string.IsNullOrEmpty(x.Sampler)) rows.Add(R("Sampler", x.Sampler));
            if (!string.IsNullOrEmpty(x.SizeText)) rows.Add(R("Size", x.SizeText));
            if (!string.IsNullOrEmpty(x.Scale)) rows.Add(R("Scale", x.Scale));

            return rows;
        }


        private Tuple<string, string> T(string a, string b) => Tuple.Create(a, b ?? "");

        // =======================================================================================
        // 既存のメタ抽出（PNG/EXIF/XMP） …ロバスト版
        // =======================================================================================
        private static Dictionary<string, object> Extract(string path)
        {
            var dict = new Dictionary<string, object>();
            dict["file"] = Path.GetFileName(path);

            var dirs = ImageMetadataReader.ReadMetadata(path);
            var raw = new Dictionary<string, object>();

            // --- PNG: tEXt / iTXt / zTXt（parameters など）
            var pngText = GetPngText(dirs);
            if (pngText.Count > 0)
            {
                raw["png"] = pngText;

                var cands = new List<string>();
                foreach (var k in new[] { "parameters", "prompt", "negative_prompt", "sd-metadata" })
                    if (pngText.TryGetValue(k, out var v)) cands.Add(v);

                // ★ 追加：Textual Data の「ブロック全文」も候補に入れる
                var blocks = GetPngTextBlocks(dirs);
                if (blocks.Count > 0) cands.AddRange(blocks);

                if (cands.Count > 0)
                    dict["parsed"] = ParseSdBlock(string.Join("\n", cands));
            }


            // --- EXIF
            var ifd0 = dirs.OfType<ExifIfd0Directory>().FirstOrDefault();
            var subIfd = dirs.OfType<ExifSubIfdDirectory>().FirstOrDefault();
            var exif = new Dictionary<string, string>();
            var v1 = GetExifStringSafe(ifd0, ExifDirectoryBase.TagImageDescription);
            if (!string.IsNullOrEmpty(v1)) exif["ImageDescription"] = v1;
            var v2 = GetExifStringSafe(ifd0, ExifDirectoryBase.TagSoftware);
            if (!string.IsNullOrEmpty(v2)) exif["Software"] = v2;
            var v3 = GetExifStringSafe(subIfd, ExifDirectoryBase.TagUserComment);
            if (!string.IsNullOrEmpty(v3)) exif["UserComment"] = v3;
            if (exif.Count > 0) raw["exif"] = exif;

            // --- XMP
            string xmp = TryGetReadableXmpText(dirs);
            if (!string.IsNullOrEmpty(xmp)) raw["xmp"] = xmp;

            // --- 補完解析（A1111系）
            if (!dict.ContainsKey("parsed"))
            {
                var cands = new List<string>();
                if (raw.ContainsKey("exif")) foreach (var kv in (Dictionary<string, string>)raw["exif"]) AddCandidateText(cands, kv.Value);
                if (!string.IsNullOrEmpty(xmp)) AddCandidateText(cands, xmp);
                var joined = string.Join("\n", cands);
                if (HasSdMarkers(joined)) dict["parsed"] = ParseSdBlock(joined);
            }

            raw["dirs"] = DumpAllDirectories(dirs);
            dict["raw"] = raw; // UI から非表示切替可
            return dict;
        }

        private static Dictionary<string, string> GetPngText(IReadOnlyList<MetadataExtractor.Directory> dirs)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var d in dirs)
            {
                if (!d.Name.StartsWith("PNG", StringComparison.OrdinalIgnoreCase))
                    continue;

                string lastKey = null;
                foreach (var tag in d.Tags)
                {
                    var name = tag.Name ?? "";
                    var desc = tag.Description ?? "";

                    if (string.Equals(name, "Keyword", StringComparison.OrdinalIgnoreCase)) lastKey = desc;
                    else if (string.Equals(name, "Text", StringComparison.OrdinalIgnoreCase) && lastKey != null)
                    { result[lastKey] = desc; lastKey = null; }
                    else if (string.Equals(name, "Textual Data", StringComparison.OrdinalIgnoreCase))
                    {
                        var lines = desc.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var line in lines)
                        {
                            var idx = line.IndexOf(':'); if (idx <= 0) continue;
                            var key = line.Substring(0, idx).Trim();
                            var val = line.Substring(idx + 1).Trim();
                            if (key.Length > 0) result[key] = val;
                        }
                    }
                }
            }
            return result;
        }

        // =======================================================================================
        // PNG の "Textual Data" の説明文ブロックを生で取得（A1111のparameters全文が入ることがあるため）
        // =======================================================================================
        private static List<string> GetPngTextBlocks(IReadOnlyList<MetadataExtractor.Directory> dirs)
        {
            var list = new List<string>();
            foreach (var d in dirs)
            {
                if (!d.Name.StartsWith("PNG", StringComparison.OrdinalIgnoreCase)) continue;
                foreach (var tag in d.Tags)
                {
                    if (string.Equals(tag.Name, "Textual Data", StringComparison.OrdinalIgnoreCase))
                    {
                        var desc = tag.Description ?? "";
                        if (!string.IsNullOrWhiteSpace(desc) && HasSdMarkers(desc))
                            list.Add(desc);
                    }
                }
            }
            return list;
        }

        // =======================================================================================
        // EXIF用 安全デコード（UserCommentの8バイト識別子対応）
        // =======================================================================================
        private static string GetExifStringSafe(ExifDirectoryBase dir, int tag)
        {
            if (dir == null) return null;

            try
            {
                var s = dir.GetString(tag);
                if (!string.IsNullOrEmpty(s) && !IsProbablyBinaryNumbers(s) && IsLikelyHumanText(s))
                    return s;
            }
            catch { }

            byte[] bytes = null;
            try { bytes = dir.GetByteArray(tag); } catch { }
            if (bytes == null || bytes.Length == 0) return null;

            if (bytes.Length >= 8)
            {
                var code = Encoding.ASCII.GetString(bytes, 0, 8);
                var body = bytes.Skip(8).ToArray();

                if (code.StartsWith("ASCII"))
                {
                    var txt = Encoding.ASCII.GetString(body).TrimEnd('\0');
                    if (IsSdOrReadable(txt)) return txt;
                }
                else if (code.StartsWith("UTF-8"))
                {
                    var txt = Encoding.UTF8.GetString(body).TrimEnd('\0');
                    if (IsSdOrReadable(txt)) return txt;
                }
                else if (code.StartsWith("UNICODE"))
                {
                    var le = Encoding.Unicode.GetString(body).TrimEnd('\0');
                    if (IsSdOrReadable(le)) return le;
                    var be = Encoding.BigEndianUnicode.GetString(body).TrimEnd('\0');
                    if (IsSdOrReadable(be)) return be;
                }
            }

            try
            {
                var u16 = Encoding.Unicode.GetString(bytes).TrimEnd('\0');
                if (IsSdOrReadable(u16)) return u16;
            }
            catch { }

            var cands = new List<string>();
            try { cands.Add(Encoding.UTF8.GetString(bytes).TrimEnd('\0')); } catch { }
            try { cands.Add(Encoding.ASCII.GetString(bytes).TrimEnd('\0')); } catch { }
            try { cands.Add(Encoding.BigEndianUnicode.GetString(bytes).TrimEnd('\0')); } catch { }

            foreach (var c in cands) if (HasSdMarkers(c)) return c;
            foreach (var c in cands) if (IsLikelyHumanText(c) && !IsProbablyBinaryNumbers(c)) return c;

            return null;
        }
        private static bool IsSdOrReadable(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            if (HasSdMarkers(s)) return true;
            return IsLikelyHumanText(s) && !IsProbablyBinaryNumbers(s);
        }
        private static double PrintableScore(string s)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            int printable = s.Count(ch => ch >= 0x20 && ch != 0x7F);
            return printable / (double)s.Length;
        }
        private static bool IsLikelyHumanText(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            double pr = PrintableScore(s);
            int asciiLetters = s.Count(ch => ch <= 0x7E && ch >= 0x20);
            double ar = asciiLetters / (double)Math.Max(1, s.Length);
            return (pr >= 0.6 && ar >= 0.3) || HasSdMarkers(s);
        }
        private static string TryGetReadableXmpText(IReadOnlyList<MetadataExtractor.Directory> dirs)
        {
            var xmpDir = dirs.FirstOrDefault(d => d.Name == "XMP");
            if (xmpDir == null) return null;

            var sb = new StringBuilder();
            foreach (var t in xmpDir.Tags)
            {
                var name = t.Name ?? "";
                var desc = t.Description ?? "";
                if (IsProbablyBinaryNumbers(desc)) continue;
                if (!IsLikelyHumanText(desc) && !HasSdMarkers(desc)) continue;
                sb.AppendLine($"{name}: {desc}");
            }
            var s = sb.ToString().Trim();
            if (string.IsNullOrEmpty(s)) return null;
            return s;
        }
        private static bool IsProbablyBinaryNumbers(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            int digitsOrSep = s.Count(c => char.IsDigit(c) || char.IsWhiteSpace(c) || c == ',' || c == ';');
            int letters = s.Count(c => char.IsLetter(c));
            double ratio = digitsOrSep / (double)Math.Max(1, s.Length);
            return ratio > 0.88 && letters < Math.Max(5, s.Length / 50);
        }
        private static bool HasSdMarkers(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            if (Regex.IsMatch(s, "(?i)(Negative\\s*prompt:|Steps\\s*:|Sampler\\s*:|CFG\\s*scale\\s*:|Seed\\s*:|Model\\s*:|Checkpoint\\s*:|Hires\\s*(steps|upscale|upscaler)\\s*:)")
                || Regex.IsMatch(s, "(?i)\"(prompt|negative[_ ]?prompt|sampler|steps|cfg(_| )?scale|seed|model)\"\\s*:"))
                return true;
            return false;
        }
        private static void AddCandidateText(List<string> list, string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return;
            if (IsProbablyBinaryNumbers(s)) return;
            if (!HasSdMarkers(s)) return;
            list.Add(s);
        }
        private static Dictionary<string, string> ParseSdBlock(string txt)
        {
            var result = new Dictionary<string, string>();
            if (!HasSdMarkers(txt)) return result;

            var negIdx = Regex.Match(txt, "(?i)Negative prompt:\\s*");
            if (negIdx.Success)
            {
                var head = txt.Substring(0, negIdx.Index).Trim();
                var tail = txt.Substring(negIdx.Index + negIdx.Length);

                // ★ 追加：先頭の "parameters:" や "prompt:" ラベルを外す
                head = Regex.Replace(head, @"(?i)^\s*(parameters|prompt)\s*:\s*", "");

                result["prompt"] = head;

                var stop = Regex.Match(tail, "\\n(?:Steps|Sampler|CFG\\s*scale|Seed|Model|Size|Version|Hires|Denoising)\\s*:", RegexOptions.IgnoreCase);
                result["negative_prompt"] = stop.Success ? tail.Substring(0, stop.Index).Trim() : tail.Trim();
            }
            else
            {
                // ラベル除去も適用
                var head = Regex.Replace(txt.Trim(), @"(?i)^\s*(parameters|prompt)\s*:\s*", "");
                result["prompt"] = head;
            }

            string[] keys = {
                                "Steps","Sampler","CFG scale","CFG Scale","cfgScale","Seed","Model","Model hash",
                                "Clip skip","Hires steps","Hires upscale","Hires upscaler","Denoising strength",
                                "Size","Version","Checkpoint","Checkpoint Name"
                            };
            foreach (var k in keys)
            {
               　//var m = Regex.Match(txt, $"{Regex.Escape(k)}\\s*:\\s*([^\\r\\n]+)", RegexOptions.IgnoreCase);
                var m = Regex.Match(txt, $"{Regex.Escape(k)}\\s*:\\s*([^,\\r\\n;]+)", RegexOptions.IgnoreCase);
                if (m.Success) result[k.ToLower().Replace(" ", "_")] = m.Groups[1].Value.Trim();
            }

            if (!result.ContainsKey("size"))
            {
                var m = Regex.Match(txt, "(\\d+)\\s*[xX]\\s*(\\d+)");
                if (m.Success) result["size"] = $"{m.Groups[1].Value}x{m.Groups[2].Value}";
            }
            return result;
        }

        // =======================================================================================
        // A1111系の <lora:name:0.XX> / <lyco:name:0.XX> を抽出
        // =======================================================================================
        private static List<LoraInfo> ExtractLorasFromPrompt(string s)
        {
            var list = new List<LoraInfo>();
            if (string.IsNullOrEmpty(s)) return list;

            var rx = new Regex(@"<\s*(lora|lyco)\s*:\s*([^:>]+)(?::\s*([0-9]*\.?[0-9]+))?\s*>",
                               RegexOptions.IgnoreCase);
            foreach (Match m in rx.Matches(s))
            {
                var name = m.Groups[2].Value.Trim();
                var w = m.Groups[3].Success ? m.Groups[3].Value : null;
                list.Add(new LoraInfo { Name = name, Weight = w });
            }
            return list;
        }

        // =======================================================================================
        // civitai:288584@324619 → https://civitai.com/models/288584?modelVersionId=324619
        // civitai:288584        → https://civitai.com/models/288584
        // それ以外（名前のみ等）  → https://civitai.com/?query=<URLENCODE(name)>
        // =======================================================================================
        // 引数3つ版
        // 既存（呼び出し側の後方互換用）
        private bool TryMakeCivitaiUrl(string text, string hash, out string url)
            => TryMakeCivitaiUrl(text, hash, null, out url);

        // 新規（VersionId 優先対応）
        private bool TryMakeCivitaiUrl(string text, string hash, string versionId, out string url)
        {
            url = null;
            var s = TrimUrn(text);

            // 1) URN civitai:modelId@versionId
            var m = Regex.Match(s, @"civitai:(\d+)(?:@(\d+))?");
            if (m.Success)
            {
                var mid = m.Groups[1].Value;
                var vid = m.Groups[2].Success ? m.Groups[2].Value : null;
                url = (vid != null)
                    ? $"https://civitai.com/models/{mid}?modelVersionId={vid}"
                    : $"https://civitai.com/models/{mid}";
                return true;
            }

            // 2) Hash 解決（成功＝直リンク）
            if (!string.IsNullOrEmpty(hash) && TryResolveCivitaiByHash(hash, out url)) return true;

            // 3) VersionId 解決（成功＝直リンク）
            if (!string.IsNullOrEmpty(versionId) && TryResolveCivitaiByVersionId(versionId, out url)) return true;

            // 4) 検索へフォールバック
            var q = !string.IsNullOrEmpty(hash) ? hash : s;
            url = "https://civitai.com/?query=" + Uri.EscapeDataString(q);
            return true;
        }

        // =======================================================================================
        // raw["png"]["Steps"] の中にある `Lora hashes: "name: hash, name2: hash2, ..."` を辞書化する処理
        // =======================================================================================
        private static Dictionary<string, string> ParseLoraHashesFromSteps(string stepsLine)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(stepsLine)) return map;

            var m = System.Text.RegularExpressions.Regex.Match(stepsLine, @"Lora hashes:\s*""([^""]+)""", RegexOptions.IgnoreCase);
            if (!m.Success) return map;

            var body = m.Groups[1].Value; 
            foreach (var part in body.Split(','))
            {
                var kv = part.Split(new[] { ':' }, 2);
                if (kv.Length != 2) continue;
                var name = kv[0].Trim();
                var hash = kv[1].Trim();
                if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(hash))
                    map[name] = hash;
            }
            return map;
        }

        private static Dictionary<string, string> ParseTiHashesFromSteps(string stepsLine)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(stepsLine)) return map;

            var m = Regex.Match(stepsLine, @"TI hashes:\s*""([^""]+)""", RegexOptions.IgnoreCase);
            if (!m.Success) return map;

            var body = m.Groups[1].Value;
            foreach (var part in body.Split(','))
            {
                var kv = part.Split(new[] { ':' }, 2);
                if (kv.Length != 2) continue;
                var name = kv[0].Trim();
                var hash = kv[1].Trim();
                if (name.Length > 0 && hash.Length > 0) map[name] = hash;
            }
            return map;
        }


        // Civitai の by-hash API を使って、ハッシュから modelId / modelVersionId を解決
        // 例: https://civitai.com/api/v1/model-versions/by-hash/939a01b898c5
        private bool TryResolveCivitaiByHash(string hash, out string directUrl)
        {
            directUrl = null;
            if (string.IsNullOrWhiteSpace(hash)) return false;

            try
            {
                // .NET 4.8 で TLS まわりの既定を安全側に
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

                using (var http = new HttpClient())
                {
                    http.Timeout = TimeSpan.FromSeconds(6);
                    http.DefaultRequestHeaders.UserAgent.ParseAdd("Meta_DataSearch/1.0");

                    var resp = http.GetAsync("https://civitai.com/api/v1/model-versions/by-hash/" + hash).Result;
                    if (!resp.IsSuccessStatusCode) return false;

                    var json = resp.Content.ReadAsStringAsync().Result;
                    var jo = JsonConvert.DeserializeObject<JObject>(json);
                    if (jo == null) return false;

                    var modelId = jo["modelId"]?.ToString();
                    var verId = jo["id"]?.ToString();  // modelVersionId
                    if (!string.IsNullOrEmpty(modelId) && !string.IsNullOrEmpty(verId))
                    {
                        directUrl = $"https://civitai.com/models/{modelId}?modelVersionId={verId}";
                        return true;
                    }
                }
            }
            catch { /* 失敗時は検索へフォールバック */ }

            return false;
        }

        // キャッシュ（同一版IDでの連続アクセスを抑制）
        private static readonly Dictionary<string, string> _verUrlCache =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private bool TryResolveCivitaiByVersionId(string versionId, out string directUrl)
        {
            directUrl = null;
            if (string.IsNullOrWhiteSpace(versionId)) return false;

            string cached;
            if (_verUrlCache.TryGetValue(versionId, out cached)) { directUrl = cached; return true; }

            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                using (var http = new HttpClient())
                {
                    http.Timeout = TimeSpan.FromSeconds(6);
                    http.DefaultRequestHeaders.UserAgent.ParseAdd("Meta_DataSearch/1.0");
                    var resp = http.GetAsync("https://civitai.com/api/v1/model-versions/" + versionId).Result;
                    if (!resp.IsSuccessStatusCode) return false;
                    var json = resp.Content.ReadAsStringAsync().Result;

                    var jo = JsonConvert.DeserializeObject<JObject>(json);
                    var modelId = jo?["modelId"]?.ToString();
                    if (!string.IsNullOrEmpty(modelId))
                    {
                        directUrl = $"https://civitai.com/models/{modelId}?modelVersionId={versionId}";
                        _verUrlCache[versionId] = directUrl;
                        return true;
                    }
                }
            }
            catch { /* ネットワーク不可なら false */ }

            return false;
        }


        // =======================================================================================
        // .parameters テキストから Checkpoint と LoRA を抽出
        // =======================================================================================
        private void ExtractFromA1111Parameters(string parameters, SdInfo info)
        {
            if (string.IsNullOrWhiteSpace(parameters)) return;

            // Checkpoint (Model / Model hash)
            var mModel = Regex.Match(parameters, @"(^|\n)\s*Model:\s*([^\r\n]+)", RegexOptions.IgnoreCase);
            var mHash = Regex.Match(parameters, @"(^|\n)\s*Model hash:\s*([0-9a-fA-F]+)", RegexOptions.IgnoreCase);
            if (mModel.Success && string.IsNullOrEmpty(info.Checkpoint))
            {
                var name = mModel.Groups[2].Value.Trim();
                var hash = mHash.Success ? mHash.Groups[2].Value.Trim() : null;
                info.Checkpoint = string.IsNullOrEmpty(hash) ? name : $"{name} ({hash})";
                if (string.IsNullOrEmpty(info.ModelHash) && mHash.Success)
                    info.ModelHash = hash;
            }

            // <lora:name:0.8[:clip]>
            foreach (Match m in Regex.Matches(parameters, @"<lora:([^:>]+)(?::([0-9.]+))?(?::([0-9.]+))?>",
                     RegexOptions.IgnoreCase))
            {
                AddLoraUnique(info, m.Groups[1].Value.Trim(),
                                   string.IsNullOrEmpty(m.Groups[2].Value) ? null : m.Groups[2].Value);
            }

            // LoRA: Name [, weight: 0.8] / LyCORIS: Name [, weight: 0.6]
            foreach (Match m in Regex.Matches(parameters,
                     @"(^|\n)\s*(LoRA|LyCORIS)\s*:\s*([^,\r\n]+)(?:[^0-9\r\n]*\bweight\s*:?\s*([0-9.]+))?",
                     RegexOptions.IgnoreCase))
            {
                var name = m.Groups[3].Value.Trim();
                var w = m.Groups[4].Success ? m.Groups[4].Value.Trim() : null;
                if (!string.IsNullOrEmpty(name)) AddLoraUnique(info, name, w);
            }

            // <embedding:name:0.8> 形式
            foreach (Match m in Regex.Matches(parameters, @"<embedding:([^:>]+)(?::([0-9.]+))?>",
                     RegexOptions.IgnoreCase))
            {
                var name = m.Groups[1].Value.Trim();
                var w = m.Groups[2].Success ? m.Groups[2].Value.Trim() : null;
                AddEmbeddingUnique(info, name, w);
            }

            // AddNet Model / Weight （順番ベースでペアリング）
            var models = Regex.Matches(parameters, @"(^|\n)AddNet Model \d+:\s*(.+)", RegexOptions.IgnoreCase);
            var weights = Regex.Matches(parameters, @"(^|\n)AddNet Weight \d+:\s*([0-9.]+)", RegexOptions.IgnoreCase);
            for (int i = 0; i < models.Count; i++)
            {
                var name = models[i].Groups[2].Value.Trim();
                string w = (i < weights.Count) ? weights[i].Groups[2].Value.Trim() : null;
                if (!string.IsNullOrEmpty(name)) AddLoraUnique(info, name, w);
            }

            // Lora hashes: "name: hash, ..."
            var mh = Regex.Match(parameters, @"Lora hashes:\s*""([^""]+)""", RegexOptions.IgnoreCase);
            if (mh.Success)
            {
                var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var part in mh.Groups[1].Value.Split(','))
                {
                    var kv = part.Split(new[] { ':' }, 2);
                    if (kv.Length == 2) dict[kv[0].Trim()] = kv[1].Trim();
                }
                // 既存LoRAへハッシュ付与
                foreach (var l in info.Loras)
                {
                    string h; if (l.Hash == null && dict.TryGetValue(l.Name, out h)) l.Hash = h;
                }
                // 無かった名前があれば行を起こしておく
                if (info.Loras.Count == 0)
                    foreach (var kv in dict) AddLoraUnique(info, kv.Key, null);
            }
            // LyCORIS hashes: "name: hash, ..."
            var lyh = Regex.Match(parameters, @"LyCORIS hashes:\s*""([^""]+)""", RegexOptions.IgnoreCase);
            if (lyh.Success)
            {
                var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var part in lyh.Groups[1].Value.Split(','))
                {
                    var kv = part.Split(new[] { ':' }, 2);
                    if (kv.Length == 2) dict[kv[0].Trim()] = kv[1].Trim();
                }
                foreach (var l in info.Loras)
                {
                    string h; if (l.Hash == null && dict.TryGetValue(l.Name, out h)) l.Hash = h;
                }
                if (info.Loras.Count == 0)
                    foreach (var kv in dict) AddLoraUnique(info, kv.Key, null);
            }
            // TI hashes: "name: hash, name2: hash2, ..."
            var tih = Regex.Match(parameters, @"TI hashes:\s*""([^""]+)""", RegexOptions.IgnoreCase);
            if (tih.Success)
            {
                var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var body = tih.Groups[1].Value;
                foreach (var part in body.Split(','))
                {
                    var kv = part.Split(new[] { ':' }, 2);
                    if (kv.Length != 2) continue;
                    var name = kv[0].Trim();
                    var hash = kv[1].Trim();
                    if (name.Length > 0 && hash.Length > 0) dict[name] = hash;
                }

                // 既存 Embedding にハッシュ付与
                foreach (var e in info.Embeddings)
                {
                    string h;
                    if (e.Hash == null && dict.TryGetValue(e.Name, out h))
                        e.Hash = h;
                }
                // まだ Embedding 行がなければ、TI hashes から起こす
                if (info.Embeddings.Count == 0)
                    foreach (var kv in dict) AddEmbeddingUnique(info, kv.Key, null);
            }

        }

        // =======================================================================================
        // 全メタデータを抽出する処理（Comfy JSON と A1111でカバーできないデータを確認する）
        // =======================================================================================
        private string DumpAllMetadataToJson(string imagePath)
        {
            var root = new Dictionary<string, object>();
            root["file"] = Path.GetFileName(imagePath);

            var dirs = ImageMetadataReader.ReadMetadata(imagePath);
            var map = new Dictionary<string, object>();
            foreach (var d in dirs)
            {
                var dname = d.Name;
                var list = new List<object>();
                foreach (var t in d.Tags)
                {
                    // 値の生バイトも拾う（可能な範囲で）
                    string value;
                    try { value = t.Description ?? t.ToString(); }
                    catch { value = t.ToString(); }

                    list.Add(new Dictionary<string, object>
                    {
                        ["tag"] = t.TagName,
                        ["desc"] = value
                    });
                }
                // エラーも表示
                if (d.HasError)
                {
                    foreach (var err in d.Errors) list.Add(new Dictionary<string, object>
                    {
                        ["error"] = err
                    });
                }
                if (list.Count > 0) map[dname] = list;
            }
            root["metadata_extractor"] = map;

            // 既存の meta["raw"] を持っているならそれも併記（あなたのパイプで作っている辞書）
            // ある場合:
            // root["raw_from_reader"] = meta["raw"];

            return Newtonsoft.Json.JsonConvert.SerializeObject(root, Newtonsoft.Json.Formatting.Indented);
        }

        // =======================================================================================
        // 画像に含まれる全ディレクトリ/タグをrawDataのままダンプ（値/説明/バイト列HEX）
        // =======================================================================================
        private static Dictionary<string, object> DumpAllDirectories(IReadOnlyList<MetadataExtractor.Directory> dirs)
        {
            var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            foreach (var d in dirs)
            {
                var items = new List<Dictionary<string, object>>();
                foreach (var tag in d.Tags)
                {
                    var row = new Dictionary<string, object>();
                    row["tag"] = tag.Type;                     // 数値タグID
                    row["name"] = tag.Name;                     // タグ名
                                                                // 文字列表現（ライブラリのDescription）
                    try { if (!string.IsNullOrEmpty(tag.Description)) row["desc"] = tag.Description; } catch { }

                    // 生値（ある場合）
                    object val = null;
                    try { val = d.GetObject(tag.Type); } catch { }
                    if (val != null && !(val is byte[]))
                    {
                        row["value"] = val.ToString();
                    }
                    else if (val is byte[] bytes)
                    {
                        // HEX化（長すぎるときは切り詰め）
                        int take = Math.Min(bytes.Length, RAW_BYTES_LIMIT);
                        var sb = new StringBuilder(take * 3);
                        for (int i = 0; i < take; i++)
                        {
                            sb.Append(bytes[i].ToString("X2"));
                            if (i < take - 1) sb.Append(' ');
                        }
                        var hex = sb.ToString();
                        if (bytes.Length > take) hex += $" … (+{bytes.Length - take} bytes)";
                        row["bytes_hex"] = hex;
                        row["bytes_len"] = bytes.Length;
                    }

                    items.Add(row);
                }
                result[d.Name] = items;
            }
            return result;
        }

        // =======================================================================================
        // 画像に含まれる全ディレクトリ/タグをrawDataのままダンプ（値/説明/バイト列HEX）
        // =======================================================================================
        private void SaveRawJsonForCurrent()
        {
            if (_lastInfos == null || _lastInfos.Count == 0)
            {
                MessageBox.Show(this, "先に画像を読み込んでください。");
                return;
            }
            int saved = 0;
            for (int i = 0; i < _lastInfos.Count; i++)
            {
                var info = _lastInfos[i];
                if (string.IsNullOrEmpty(info.FullPath) || !File.Exists(info.FullPath)) continue;

                var outPath = Path.ChangeExtension(info.FullPath, ".raw.json");
                File.WriteAllText(outPath, DumpAllMetadataToJson(info.FullPath), new UTF8Encoding(false));
                saved++;
            }
            MessageBox.Show(this, $"{saved} 件の raw JSON を保存しました。");
        }

        // =======================================================================================
        // previewに画像をサムネイルとして表示させる
        // =======================================================================================
        private static Bitmap LoadBitmapNoLock(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var img = Image.FromStream(fs, useEmbeddedColorManagement: false, validateImageData: false))
                return new Bitmap(img);
        }

        // メモリ節約のため、最大辺を制限した縮小版を生成
        private static Bitmap MakeThumbnail(Image src, int maxEdge)
        {
            var w = src.Width; var h = src.Height;
            var scale = Math.Min(1.0, (double)maxEdge / Math.Max(w, h));
            int nw = Math.Max(1, (int)Math.Round(w * scale));
            int nh = Math.Max(1, (int)Math.Round(h * scale));
            var bmp = new Bitmap(nw, nh, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.DrawImage(src, new Rectangle(0, 0, nw, nh));
            }
            return bmp;
        }
        private void UpdatePreview(string imagePath)
        {
            try
            {
                // 既存画像の開放（ファイルロック/メモリ解放）
                var old = _picPreview.Image;
                _picPreview.Image = null;
                if (old != null) old.Dispose();

                if (string.IsNullOrEmpty(imagePath) || !File.Exists(imagePath)) return;

                // WebP など System.Drawing 非対応形式は try/catch で握る
                using (var full = LoadBitmapNoLock(imagePath))
                {
                    // プレビュー枠より少し大きめを作っておくとZoomで綺麗
                    int max = Math.Max(_picPreview.Width, _picPreview.Height) * 2;
                    _picPreview.Image = MakeThumbnail(full, max);
                }
            }
            catch
            {
                // 読めない形式(例: WebP)は無視してプレビュー無し
                _picPreview.Image = null;
            }
        }

        // =======================================================================================
        // メタデータが空か確認する
        // =======================================================================================
        private static bool IsInfoEmpty(SdInfo x)
        {
            if (x == null) return true;
            bool noModel = string.IsNullOrWhiteSpace(x.Checkpoint) && x.Loras.Count == 0 && x.Embeddings.Count == 0;
            bool noTexts = string.IsNullOrWhiteSpace(x.Prompt) && string.IsNullOrWhiteSpace(x.NegativePrompt);
            bool noParams = string.IsNullOrWhiteSpace(x.Steps) && string.IsNullOrWhiteSpace(x.Sampler) &&
                            string.IsNullOrWhiteSpace(x.Cfg) && string.IsNullOrWhiteSpace(x.Seed);
            return noModel && noTexts && noParams;
        }

        private async Task PostProcessMissingMetaAsync()
        {
            // 空メタの行インデックスを収集
            var need = _lastInfos
                .Select((x, i) => new { x, i })
                .Where(t => IsInfoEmpty(t.x) && File.Exists(t.x.FullPath))
                .Select(t => t.i)
                .ToList();

            if (need.Count == 0) return;

            var msg = (need.Count == 1)
                ? $"この画像はメタデータが見つかりませんでした。\nAI でプロンプトを推定しますか？\n\n{Path.GetFileName(_lastInfos[need[0]].FullPath)}"
                : $"{need.Count} 枚の画像でメタデータが見つかりませんでした。\nAI でプロンプトを推定しますか？";

            if (MessageBox.Show(this, msg, "AI プロンプト生成", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            // 進行中カーソル
            var old = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                foreach (var idx in need)
                {
                    var res = await RunCaptionerAsync(_lastInfos[idx].FullPath);
                    var promptText = !string.IsNullOrWhiteSpace(res.tags) ? res.tags : res.caption;
                    if (!string.IsNullOrWhiteSpace(promptText))
                    {
                        _lastInfos[idx].Prompt = promptText.Trim();
                    }

                }
            }
            finally
            {
                Cursor.Current = old;
            }

            // 画面更新
            if (chkGridView.Checked) RenderGrid();
            else RenderText();
        }


        // =======================================================================================
        // Prompt分析ツールとの連携
        // =======================================================================================
        class CaptionResult { public string caption; public string tags; }

        private async Task<CaptionResult> RunCaptionerAsync(string imagePath, int timeoutMs = 120000)
        {
            var exeDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "captioner");
            var exe = Path.Combine(exeDir, "captioner.exe");
            if (!File.Exists(exe)) throw new FileNotFoundException("captioner.exe が見つかりません。", exe);

            var psi = new ProcessStartInfo
            {
                FileName = exe,
                Arguments = "\"" + imagePath + "\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WorkingDirectory = exeDir,
            };
            psi.EnvironmentVariables["BLIP_MODEL_DIR"] =
                Path.Combine(exeDir, "blip-image-captioning-base");

            using (var p = new Process { StartInfo = psi })
            {
                p.Start();
                var cts = new CancellationTokenSource(timeoutMs);
                var readOut = p.StandardOutput.ReadToEndAsync();
                var readErr = p.StandardError.ReadToEndAsync();
                await Task.WhenAny(Task.Run(() => p.WaitForExit(), cts.Token), Task.Delay(timeoutMs, cts.Token));
                if (!p.HasExited) { try { p.Kill(); } catch { } throw new TimeoutException("captioner timeout"); }

                var output = await readOut;
                var err = await readErr;

                try
                {
                    var jo = Newtonsoft.Json.Linq.JObject.Parse(output);
                    return new CaptionResult
                    {
                        caption = (string)jo["caption"],
                        tags = (string)jo["tags"]
                    };
                }
                catch
                {
                    throw new Exception("captioner failed: " + err);
                }
            }
        }

        // =======================================================================================
        // プログレスバー汎用処理
        // =======================================================================================
        private void SetBusy(bool on, string msg = null)
        {
            if (InvokeRequired) { BeginInvoke(new Action(() => SetBusy(on, msg))); return; }

            ProgressBar1.Visible = on;
            ProgressBar1.MarqueeAnimationSpeed = on ? 30 : 0;

            // 任意：UI操作を抑止したい場合
            btnOpen.Enabled = !on;
            btnSaveRaw.Enabled = !on;
            chkGridView.Enabled = !on;

        }

    }
}
