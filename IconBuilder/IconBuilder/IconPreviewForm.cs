using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Linq;
using System.Windows.Forms;

namespace IconBuilder
{
    // Modal preview shown after the user clicks "Save Selection as Icon" but before the file is
    // written. Lets the user verify the per-size pixel rendering the .ico will contain. The
    // dialog is view-only — it does not modify the supplied bitmaps. Caller owns disposal of the
    // bitmaps; the form only displays them.
    //
    // The visual design borrows the flat, light-surface aesthetic of Windows 11 settings: white
    // content surfaces over a soft grey window background, a subtle 1 px divider between the
    // sidebar and the preview, accent-blue primary button, and Segoe UI throughout.
    internal sealed class IconPreviewForm : Form
    {
        private static readonly Color AccentColor = Color.FromArgb(0, 120, 212);
        private static readonly Color AccentHover = Color.FromArgb(16, 110, 190);
        private static readonly Color AccentPressed = Color.FromArgb(0, 95, 170);
        private static readonly Color WindowBg = Color.FromArgb(243, 243, 243);
        private static readonly Color SurfaceBg = Color.White;
        private static readonly Color DividerColor = Color.FromArgb(228, 228, 228);
        private static readonly Color TextPrimary = Color.FromArgb(28, 28, 28);
        private static readonly Color TextSecondary = Color.FromArgb(100, 100, 100);
        private static readonly Color SelectionFill = Color.FromArgb(232, 240, 252);
        private static readonly Color SelectionBorder = Color.FromArgb(180, 209, 245);

        private readonly IDictionary<int, Bitmap> _sources;
        private readonly SizeListBox _sizeList;
        private readonly PreviewPanel _previewPanel;

        public IconPreviewForm(IDictionary<int, Bitmap> sources)
        {
            _sources = sources ?? throw new ArgumentNullException(nameof(sources));

            Text = "Preview Icon";
            ShowIcon = true;
            try
            {
                using (var s = System.Reflection.Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream("IconBuilder.icon-builder.ico"))
                {
                    if (s != null) Icon = new Icon(s);
                }
            }
            catch { /* non-fatal */ }
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(760, 540);
            Size = new Size(960, 660);
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = false;
            BackColor = WindowBg;
            Font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
            DoubleBuffered = true;

            // ---- Bottom button strip --------------------------------------------------------
            // NOTE: Padding is intentionally zero. With a non-zero Padding on this panel, the
            // Dock=Top divider gets pushed down by the top padding (so the line drifts away
            // from the content above and ends up right against the buttons). Keep padding at 0
            // and lay out the buttons by hand in the Resize handler.
            var buttonStrip = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 64,
                BackColor = WindowBg,
                Padding = new Padding(0),
            };
            var divider = new Panel
            {
                Dock = DockStyle.Top,
                Height = 1,
                BackColor = DividerColor,
            };
            buttonStrip.Controls.Add(divider);

            var saveBtn = new FlatAccentButton
            {
                Text = "Save",
                DialogResult = DialogResult.OK,
                Anchor = AnchorStyles.Right | AnchorStyles.Top,
                Size = new Size(90, 28),
                IsPrimary = true,
            };
            var cancelBtn = new FlatAccentButton
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Anchor = AnchorStyles.Right | AnchorStyles.Top,
                Size = new Size(90, 28),
                IsPrimary = false,
            };
            buttonStrip.Controls.Add(saveBtn);
            buttonStrip.Controls.Add(cancelBtn);
            buttonStrip.Resize += (s, e) =>
            {
                int right = buttonStrip.ClientSize.Width - 20;
                int top = 20; // 1 px divider + 19 px breathing room above the buttons
                saveBtn.Location = new Point(right - saveBtn.Width, top);
                cancelBtn.Location = new Point(saveBtn.Left - cancelBtn.Width - 8, top);
            };
            AcceptButton = saveBtn;
            CancelButton = cancelBtn;

            // ---- Left sidebar (sizes list) -------------------------------------------------
            const int SidebarWidth = 220;

            var sidebar = new Panel
            {
                Dock = DockStyle.Left,
                Width = SidebarWidth,
                BackColor = SurfaceBg,
                Padding = new Padding(0),
            };
            var sidebarRightDivider = new Panel
            {
                Dock = DockStyle.Right,
                Width = 1,
                BackColor = DividerColor,
            };
            var sidebarHeader = new Label
            {
                Text = "Sizes",
                Dock = DockStyle.Top,
                Height = 44,
                Padding = new Padding(20, 14, 20, 6),
                Font = new Font("Segoe UI Semibold", 10.5f, FontStyle.Regular, GraphicsUnit.Point),
                ForeColor = TextPrimary,
                BackColor = SurfaceBg,
                TextAlign = ContentAlignment.MiddleLeft,
            };
            _sizeList = new SizeListBox
            {
                Dock = DockStyle.Fill,
                BackColor = SurfaceBg,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular, GraphicsUnit.Point),
                IntegralHeight = false,
                ItemHeight = 44,
                DrawMode = DrawMode.OwnerDrawFixed,
            };
            sidebar.Controls.Add(_sizeList);
            sidebar.Controls.Add(sidebarHeader);
            sidebar.Controls.Add(sidebarRightDivider);

            // ---- Right side: preview -------------------------------------------------------
            var previewHost = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = WindowBg,
                Padding = new Padding(20, 16, 20, 16),
            };
            _previewPanel = new PreviewPanel
            {
                Dock = DockStyle.Fill,
                BackColor = WindowBg,
            };
            previewHost.Controls.Add(_previewPanel);

            Controls.Add(previewHost);
            Controls.Add(sidebar);
            Controls.Add(buttonStrip);

            // ---- Populate size list --------------------------------------------------------
            var orderedSizes = _sources.Keys.OrderByDescending(s => s).ToList();
            foreach (int size in orderedSizes)
            {
                _sizeList.Items.Add(new SizeItem(size, _sources[size]));
            }
            _sizeList.SelectedIndexChanged += (s, e) =>
            {
                if (_sizeList.SelectedItem is SizeItem item
                    && _sources.TryGetValue(item.Size, out Bitmap bmp))
                {
                    _previewPanel.SetPreview(bmp, item.Size);
                }
            };
            if (_sizeList.Items.Count > 0) _sizeList.SelectedIndex = 0;
        }

        // ---------- Owner-drawn list item ----------------------------------------------------

        private sealed class SizeItem
        {
            public int Size { get; }
            public Bitmap Thumbnail { get; }
            public SizeItem(int size, Bitmap thumb) { Size = size; Thumbnail = thumb; }
        }

        // List with hover/selection highlight and a small thumbnail chip per item. Keeps the
        // sidebar visually closer to a Windows 11 settings nav list than a plain ListBox would.
        private sealed class SizeListBox : ListBox
        {
            private int _hoverIndex = -1;

            public SizeListBox()
            {
                SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            }

            protected override void OnMouseMove(MouseEventArgs e)
            {
                base.OnMouseMove(e);
                int idx = IndexFromPoint(e.Location);
                if (idx != _hoverIndex)
                {
                    int prev = _hoverIndex;
                    _hoverIndex = idx;
                    if (prev >= 0) Invalidate(GetItemRectangle(prev));
                    if (idx >= 0) Invalidate(GetItemRectangle(idx));
                }
            }

            protected override void OnMouseLeave(EventArgs e)
            {
                base.OnMouseLeave(e);
                if (_hoverIndex >= 0)
                {
                    int prev = _hoverIndex;
                    _hoverIndex = -1;
                    if (prev < Items.Count) Invalidate(GetItemRectangle(prev));
                }
            }

            protected override void OnDrawItem(DrawItemEventArgs e)
            {
                if (e.Index < 0 || e.Index >= Items.Count) return;
                var item = (SizeItem)Items[e.Index];
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
                bool hover = e.Index == _hoverIndex;

                Rectangle full = e.Bounds;
                Rectangle pill = Rectangle.Inflate(full, -8, -3);

                using (var bg = new SolidBrush(SurfaceBg))
                    g.FillRectangle(bg, full);

                if (selected)
                {
                    using (var path = RoundedRect(pill, 6))
                    using (var fill = new SolidBrush(SelectionFill))
                    using (var pen = new Pen(SelectionBorder, 1f))
                    {
                        g.FillPath(fill, path);
                        g.DrawPath(pen, path);
                    }
                }
                else if (hover)
                {
                    using (var path = RoundedRect(pill, 6))
                    using (var fill = new SolidBrush(Color.FromArgb(245, 245, 245)))
                    {
                        g.FillPath(fill, path);
                    }
                }

                // Thumbnail chip on the left of the row.
                int chip = 28;
                int chipX = pill.Left + 8;
                int chipY = pill.Top + (pill.Height - chip) / 2;
                var chipRect = new Rectangle(chipX, chipY, chip, chip);
                DrawThumbnail(g, chipRect, item.Thumbnail);

                // Primary label: "32 × 32"
                string label = item.Size + " × " + item.Size;
                string sub = item.Size == IconWriter.PngSize ? "PNG" : "BMP";
                using (var labelFont = new Font("Segoe UI", 9.5f, FontStyle.Regular))
                using (var subFont = new Font("Segoe UI", 8.25f, FontStyle.Regular))
                using (var labelBrush = new SolidBrush(TextPrimary))
                using (var subBrush = new SolidBrush(TextSecondary))
                {
                    int textX = chipRect.Right + 12;
                    int labelH = (int)Math.Ceiling(labelFont.GetHeight(g));
                    int subH = (int)Math.Ceiling(subFont.GetHeight(g));
                    int totalH = labelH + subH;
                    int textY = pill.Top + (pill.Height - totalH) / 2;

                    g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                    g.DrawString(label, labelFont, labelBrush, textX, textY);
                    g.DrawString(sub, subFont, subBrush, textX, textY + labelH);
                }
            }

            private static void DrawThumbnail(Graphics g, Rectangle dest, Bitmap bmp)
            {
                using (var path = RoundedRect(dest, 4))
                {
                    Region prev = g.Clip;
                    g.SetClip(path);
                    DrawCheckerboard(g, dest, 4);
                    if (bmp != null)
                    {
                        var oldI = g.InterpolationMode;
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.DrawImage(bmp, dest);
                        g.InterpolationMode = oldI;
                    }
                    g.Clip = prev;

                    using (var pen = new Pen(Color.FromArgb(220, 220, 220), 1f))
                        g.DrawPath(pen, path);
                }
            }
        }

        // ---------- Preview canvas -----------------------------------------------------------

        private sealed class PreviewPanel : Panel
        {
            private Bitmap _bmp;
            private int _iconSize;

            public PreviewPanel()
            {
                DoubleBuffered = true;
                ResizeRedraw = true;
            }

            public void SetPreview(Bitmap bmp, int iconSize)
            {
                _bmp = bmp;
                _iconSize = iconSize;
                Invalidate();
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);

                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

                // Rounded white card filling the panel: visually contains the preview content
                // and lifts it off the grey window background.
                var card = new Rectangle(0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
                using (var path = RoundedRect(card, 8))
                using (var fill = new SolidBrush(SurfaceBg))
                using (var border = new Pen(DividerColor, 1f))
                {
                    g.FillPath(fill, path);
                    g.DrawPath(border, path);
                }

                if (_bmp == null || _iconSize <= 0) return;

                // Reserve a header strip at the top for the size caption.
                const int HeaderH = 48;
                int margin = 24;
                int contentTop = HeaderH;
                int availW = Math.Max(1, ClientSize.Width - margin * 2);
                int availH = Math.Max(1, ClientSize.Height - contentTop - margin);

                int zoom = Math.Max(1, Math.Min(availW / _iconSize, availH / _iconSize));
                int drawW = _iconSize * zoom;
                int drawH = _iconSize * zoom;
                int drawX = (ClientSize.Width - drawW) / 2;
                int drawY = contentTop + (availH - drawH) / 2;
                var drawRect = new Rectangle(drawX, drawY, drawW, drawH);

                // Subtle shadow under the icon canvas to lift it off the card surface.
                for (int i = 4; i >= 1; i--)
                {
                    var shadow = new Rectangle(drawRect.X - i, drawRect.Y - i + 2,
                        drawRect.Width + i * 2, drawRect.Height + i * 2);
                    using (var p = RoundedRect(shadow, 4))
                    using (var b = new SolidBrush(Color.FromArgb(8, 0, 0, 0)))
                        g.FillPath(b, p);
                }

                DrawCheckerboard(g, drawRect, 8);

                var oldInterp = g.InterpolationMode;
                var oldOffset = g.PixelOffsetMode;
                try
                {
                    g.InterpolationMode = InterpolationMode.NearestNeighbor;
                    g.PixelOffsetMode = PixelOffsetMode.Half;
                    g.DrawImage(_bmp, drawRect);
                }
                finally
                {
                    g.InterpolationMode = oldInterp;
                    g.PixelOffsetMode = oldOffset;
                }

                if (_iconSize != IconWriter.PngSize && zoom >= 4)
                {
                    using (var pen = new Pen(Color.FromArgb(180, 0, 0, 0), 1f))
                    {
                        for (int i = 0; i <= _iconSize; i++)
                        {
                            int x = drawX + i * zoom;
                            g.DrawLine(pen, x, drawY, x, drawY + drawH);
                            int y = drawY + i * zoom;
                            g.DrawLine(pen, drawX, y, drawX + drawW, y);
                        }
                    }
                }

                using (var pen = new Pen(Color.FromArgb(210, 210, 210), 1f))
                    g.DrawRectangle(pen, drawRect);

                // Top header text.
                string title = _iconSize + " × " + _iconSize;
                string subtitle = (_iconSize == IconWriter.PngSize ? "PNG" : "BMP")
                    + "  ·  zoom " + zoom + "×";
                using (var titleFont = new Font("Segoe UI Semibold", 12f, FontStyle.Regular))
                using (var subFont = new Font("Segoe UI", 9f, FontStyle.Regular))
                using (var titleBrush = new SolidBrush(TextPrimary))
                using (var subBrush = new SolidBrush(TextSecondary))
                {
                    g.DrawString(title, titleFont, titleBrush, 20, 14);
                    SizeF ts = g.MeasureString(title, titleFont);
                    g.DrawString(subtitle, subFont, subBrush, 20 + ts.Width + 10, 20);
                }
            }
        }

        // ---------- Modern flat button -------------------------------------------------------

        private sealed class FlatAccentButton : Button
        {
            public bool IsPrimary { get; set; }
            private bool _hover;
            private bool _pressed;

            public FlatAccentButton()
            {
                FlatStyle = FlatStyle.Flat;
                FlatAppearance.BorderSize = 0;
                Cursor = Cursors.Hand;
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);
                SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint
                    | ControlStyles.AllPaintingInWmPaint | ControlStyles.ResizeRedraw, true);
                BackColor = Color.Transparent;
            }

            protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
            protected override void OnMouseLeave(EventArgs e) { _hover = false; _pressed = false; Invalidate(); base.OnMouseLeave(e); }
            protected override void OnMouseDown(MouseEventArgs mevent) { _pressed = true; Invalidate(); base.OnMouseDown(mevent); }
            protected override void OnMouseUp(MouseEventArgs mevent) { _pressed = false; Invalidate(); base.OnMouseUp(mevent); }

            protected override void OnPaint(PaintEventArgs pevent)
            {
                var g = pevent.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

                Color fill, text;
                if (IsPrimary)
                {
                    fill = _pressed ? AccentPressed : (_hover ? AccentHover : AccentColor);
                    text = Color.White;
                }
                else
                {
                    fill = _pressed ? Color.FromArgb(225, 225, 225)
                         : (_hover ? Color.FromArgb(238, 238, 238) : Color.FromArgb(249, 249, 249));
                    text = TextPrimary;
                }

                var rect = new Rectangle(0, 0, Width - 1, Height - 1);
                using (var path = RoundedRect(rect, 4))
                using (var brush = new SolidBrush(fill))
                using (var border = new Pen(IsPrimary ? fill : Color.FromArgb(210, 210, 210), 1f))
                {
                    g.FillPath(brush, path);
                    g.DrawPath(border, path);
                }

                TextRenderer.DrawText(g, Text, Font, ClientRectangle, text,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter
                    | TextFormatFlags.SingleLine | TextFormatFlags.NoPadding);

                if (Focused && ShowFocusCues)
                {
                    var focus = Rectangle.Inflate(rect, -3, -3);
                    using (var p = new Pen(IsPrimary ? Color.White : AccentColor, 1f) { DashStyle = DashStyle.Dot })
                    using (var path = RoundedRect(focus, 3))
                        g.DrawPath(p, path);
                }
            }
        }

        // ---------- Shared helpers -----------------------------------------------------------

        private static GraphicsPath RoundedRect(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            int d = radius * 2;
            if (d <= 0 || rect.Width <= d || rect.Height <= d)
            {
                path.AddRectangle(rect);
                return path;
            }
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        private static void DrawCheckerboard(Graphics g, Rectangle rect, int cell)
        {
            using (var light = new SolidBrush(Color.White))
            using (var dark = new SolidBrush(Color.FromArgb(214, 214, 214)))
            {
                g.FillRectangle(light, rect);

                Region prev = g.Clip;
                g.SetClip(rect);
                int startX = rect.X - (rect.X % cell);
                int startY = rect.Y - (rect.Y % cell);
                for (int y = startY, row = (startY - rect.Y) / cell; y < rect.Bottom; y += cell, row++)
                {
                    for (int x = startX, col = (startX - rect.X) / cell; x < rect.Right; x += cell, col++)
                    {
                        if (((row + col) & 1) == 1)
                        {
                            g.FillRectangle(dark, x, y, cell, cell);
                        }
                    }
                }
                g.Clip = prev;
            }
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(IconPreviewForm));
            this.SuspendLayout();
            // 
            // IconPreviewForm
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "IconPreviewForm";
            this.ResumeLayout(false);

        }
    }
}
