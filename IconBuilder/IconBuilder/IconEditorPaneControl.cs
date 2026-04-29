using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace IconBuilder
{
    // UserControl hosted inside the "Icon Editor Pane" CustomTaskPane. Shows one row per
    // supported icon size. Each row has:
    //   - a preview thumbnail of the bitmap currently assigned to that size
    //   - an "Assign Selected" button that captures the current PowerPoint shape selection
    //   - a "Clear" button
    // The preview also accepts image files dragged from Windows Explorer. Bottom of the pane
    // has an "Export as Icon" button that combines all assigned slots into a single .ico file.
    public class IconEditorPaneControl : UserControl
    {
        private static readonly int[] Sizes = { 16, 20, 24, 32, 40, 48, 64, 256 };
        private const int PreviewBoxSize = 72;

        private readonly Dictionary<int, Bitmap> _slotBitmaps = new Dictionary<int, Bitmap>();
        private readonly Dictionary<int, PictureBox> _slotPreviews = new Dictionary<int, PictureBox>();
        private readonly Dictionary<int, Label> _slotStatus = new Dictionary<int, Label>();
        private static Bitmap s_checkerTile;

        // Undo stack for slot mutations (Assign / Clear / Clear All Slots / drag-drop). Each
        // entry is a snapshot of the slots affected by a single user action; popping an entry
        // restores those slots to the state they were in before the action. Bound to Ctrl+Z
        // via ProcessCmdKey so it works whenever focus is anywhere inside the pane (including
        // immediately after clicking one of our buttons, which retains focus).
        private const int MaxUndoDepth = 50;
        private readonly Stack<Dictionary<int, SlotSnapshot>> _undoStack
            = new Stack<Dictionary<int, SlotSnapshot>>();

        private sealed class SlotSnapshot
        {
            // Cloned bitmap (null if the slot was empty at the time of capture). Owned by the
            // snapshot — disposed when the snapshot is popped/discarded.
            public Bitmap Bitmap;
            public string StatusText;
            public Color StatusColor;
        }

        public IconEditorPaneControl()
        {
            BuildLayout();
        }

        // Native Dock=Top layout. Controls are added in REVERSE order so they stack
        // top-to-bottom in the UserControl's client area (Dock=Top stacks based on z-order:
        // controls added later appear above earlier ones). The UserControl has AutoScroll
        // so when total docked height > viewport height, a vertical scrollbar appears
        // automatically.
        //
        // Height-for-width: HintLabel and SlotRowControl override OnSizeChanged to set
        // their own Height based on their current Width. When that happens the framework
        // natively re-docks the rest of the stack — no parent-side OnLayout overrides, no
        // manual SetBounds, no AutoScrollPosition arithmetic.
        //
        // Vertical spacing between docked controls comes from each control's Padding
        // (Dock=Top ignores Margin, but Padding creates internal whitespace inside the
        // control so the next docked control sits visually below it with a gap).
        private Label _title;
        private HintLabel _hint;
        private readonly List<SlotRowControl> _slotRows = new List<SlotRowControl>();
        private RoundedAccentButton _exportBtn;
        private Button _clearAllBtn;

        private void BuildLayout()
        {
            this.AutoScroll = true;
            this.BackColor = SystemColors.Window;
            this.Padding = new Padding(8);
            this.Font = SystemFonts.MessageBoxFont;

            _title = new Label
            {
                Text = "Icon Editor",
                Font = new Font(this.Font, FontStyle.Bold),
                AutoSize = false,
                Dock = DockStyle.Top,
                Padding = new Padding(0, 0, 0, 4),
                Height = TextRenderer.MeasureText("Mg",
                    new Font(this.Font, FontStyle.Bold)).Height + 4,
            };

            _hint = new HintLabel
            {
                Text = "Select shape(s) on the slide, then click Assign for the target size. " +
                       "You can also drop a PNG/JPG/BMP file from Explorer onto a preview.",
                Dock = DockStyle.Top,
                ForeColor = SystemColors.GrayText,
                Padding = new Padding(0, 0, 0, 8),
            };

            _exportBtn = new RoundedAccentButton(Color.FromArgb(16, 124, 16))
            {
                Text = "Export as Icon (.ico)...",
                AutoSize = true,
            };
            _exportBtn.Click += (sender, e) => OnExportClicked();

            _clearAllBtn = new Button
            {
                Text = "Clear All Slots",
                AutoSize = true,
                UseVisualStyleBackColor = true,
            };
            _clearAllBtn.Click += (sender, e) =>
            {
                PushUndo(Sizes);
                foreach (int s in Sizes) ClearSlotInternal(s);
            };

            // Wrap each button in a Dock=Top, AutoSize panel so the button keeps its natural
            // (text-fit) width — Dock=Top on the button itself would stretch it edge-to-edge.
            var exportRow = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(0, 8, 0, 8),
            };
            exportRow.Controls.Add(_exportBtn);

            var clearRow = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(0, 0, 0, 4),
            };
            clearRow.Controls.Add(_clearAllBtn);

            // Build slot rows in size order (16, 20, ..., 256).
            foreach (int size in Sizes)
            {
                int s = size;
                var row = new SlotRowControl(s, this.Font, PreviewBoxSize)
                {
                    Dock = DockStyle.Top,
                };
                row.Preview.AllowDrop = true;
                row.Preview.DragEnter += (sender, e) => OnPreviewDragEnter(e);
                row.Preview.DragDrop += (sender, e) => OnPreviewDragDrop(s, e);
                row.AssignButton.Click += (sender, e) => OnAssignClicked(s);
                row.ClearButton.Click += (sender, e) => OnClearClicked(s);

                _slotPreviews[s] = row.Preview;
                _slotStatus[s] = row.StatusLabel;

                _slotRows.Add(row);
            }

            // Add to UserControl in REVERSE visual order: last-added Dock=Top sits on top.
            // Visual order top->bottom: title, hint, slot[0..7], export, clearAll
            // => add order: clearRow, exportRow, slot[7..0], hint, title.
            this.SuspendLayout();
            this.Controls.Add(clearRow);
            this.Controls.Add(exportRow);
            for (int i = _slotRows.Count - 1; i >= 0; i--)
            {
                this.Controls.Add(_slotRows[i]);
            }
            this.Controls.Add(_hint);
            this.Controls.Add(_title);
            this.ResumeLayout(true);
        }

        // Called by SlotRowControl after computing its MinimumSize. Slot rows are Dock=Top
        // so they fit horizontally to our client width — they never overflow on their own.
        // We need an explicit AutoScrollMinSize.Width so the horizontal scrollbar appears
        // when the pane is narrower than a slot row's minimum content width.
        internal void NotifySlotRowMinWidth(int slotMinW)
        {
            int target = slotMinW + this.Padding.Horizontal;
            if (this.AutoScrollMinSize.Width < target)
            {
                this.AutoScrollMinSize = new Size(target, this.AutoScrollMinSize.Height);
            }
        }

        // Plain Label with AutoSize=false that paints word-wrapped text via TextRenderer.
        // Self-sizing height-for-width: when its Width changes (Dock=Top stretches it to
        // parent client width), it updates its own Height to fit the wrapped text plus
        // bottom Padding. The framework natively re-docks subsequent controls.
        private sealed class HintLabel : Label
        {
            public HintLabel()
            {
                AutoSize = false;
            }

            protected override void OnSizeChanged(EventArgs e)
            {
                base.OnSizeChanged(e);
                ResizeToFit();
            }

            protected override void OnTextChanged(EventArgs e)
            {
                base.OnTextChanged(e);
                ResizeToFit();
            }

            protected override void OnFontChanged(EventArgs e)
            {
                base.OnFontChanged(e);
                ResizeToFit();
            }

            private void ResizeToFit()
            {
                int w = Width - Padding.Horizontal;
                if (w <= 0) return;
                Size m = TextRenderer.MeasureText(Text ?? string.Empty, Font,
                    new Size(w, int.MaxValue),
                    TextFormatFlags.WordBreak | TextFormatFlags.NoPadding);
                int desired = Math.Max(16, m.Height + 2) + Padding.Vertical;
                if (Height != desired) Height = desired;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                var flags = TextFormatFlags.WordBreak | TextFormatFlags.Top
                    | TextFormatFlags.Left | TextFormatFlags.NoPadding;
                Rectangle r = new Rectangle(
                    Padding.Left, Padding.Top,
                    Width - Padding.Horizontal,
                    Height - Padding.Vertical);
                TextRenderer.DrawText(e.Graphics, Text, Font, r, ForeColor, flags);
            }
        }

        // (UpdateAutoScrollMinSize / FindLargestRowMinWidth removed — LayoutContent now sets
        // AutoScrollMinSize directly using the computed total content size.)

        private static Bitmap GetCheckerTile()
        {
            if (s_checkerTile != null) return s_checkerTile;
            var bmp = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                using (var br = new SolidBrush(Color.FromArgb(220, 220, 220)))
                {
                    g.FillRectangle(br, 0, 0, 8, 8);
                    g.FillRectangle(br, 8, 8, 8, 8);
                }
            }
            s_checkerTile = bmp;
            return bmp;
        }

        private static bool IsImageFile(string path)
        {
            string ext = Path.GetExtension(path)?.ToLowerInvariant();
            return ext == ".png" || ext == ".jpg" || ext == ".jpeg"
                || ext == ".bmp" || ext == ".gif" || ext == ".tif" || ext == ".tiff";
        }

        private void OnPreviewDragEnter(DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (var p in paths)
                {
                    if (IsImageFile(p)) { e.Effect = DragDropEffects.Copy; return; }
                }
            }
            e.Effect = DragDropEffects.None;
        }

        private void OnPreviewDragDrop(int size, DragEventArgs e)
        {
            try
            {
                var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (var p in paths)
                {
                    if (!IsImageFile(p)) continue;
                    using (var src = Image.FromFile(p))
                    {
                        PushUndo(new[] { size });
                        AssignBitmapForSize(size, src, "from " + Path.GetFileName(p));
                    }
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load image:\n" + ex.Message, "Icon Editor",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnAssignClicked(int size)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = null;
                try { sel = app.ActiveWindow?.Selection; } catch { sel = null; }

                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Select a shape, picture, or group on the slide first.",
                        "Icon Editor", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var range = sel.ShapeRange;
                if (range == null || range.Count == 0)
                {
                    MessageBox.Show("No shape selected.", "Icon Editor",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string sourceLabel = range.Count == 1
                    ? "from \"" + range[1].Name + "\""
                    : "from " + range.Count + " shapes";

                string tempPng = Path.Combine(Path.GetTempPath(),
                    "IconEditor_" + Guid.NewGuid().ToString("N") + ".png");
                try
                {
                    IconBuilderRibbon.ExportShapeRangeToPng(range, tempPng);
                    using (var src = Image.FromFile(tempPng))
                    {
                        PushUndo(new[] { size });
                        AssignBitmapForSize(size, src, sourceLabel);
                    }
                }
                finally
                {
                    try { if (File.Exists(tempPng)) File.Delete(tempPng); } catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to assign shape:\n" + ex.Message, "Icon Editor",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AssignBitmapForSize(int size, Image source, string statusText)
        {
            using (var trimmed = IconWriter.TrimTransparentBorder(source))
            {
                Bitmap rendered = IconWriter.ResizeTo32bpp(trimmed, size);
                SetSlotBitmap(size, rendered, statusText);
            }
        }

        private void SetSlotBitmap(int size, Bitmap bmp, string statusText)
        {
            if (_slotBitmaps.TryGetValue(size, out Bitmap old))
            {
                old?.Dispose();
            }
            _slotBitmaps[size] = bmp;

            if (_slotPreviews.TryGetValue(size, out PictureBox pb))
            {
                Image previous = pb.Image;
                pb.Image = (Bitmap)bmp.Clone();
                previous?.Dispose();
            }
            if (_slotStatus.TryGetValue(size, out Label lbl))
            {
                lbl.Text = statusText;
                lbl.ForeColor = SystemColors.ControlText;
            }
        }

        private void OnClearClicked(int size)
        {
            PushUndo(new[] { size });
            ClearSlotInternal(size);
        }

        private void ClearSlotInternal(int size)
        {
            if (_slotBitmaps.TryGetValue(size, out Bitmap old))
            {
                old?.Dispose();
                _slotBitmaps.Remove(size);
            }
            if (_slotPreviews.TryGetValue(size, out PictureBox pb))
            {
                Image previous = pb.Image;
                pb.Image = null;
                previous?.Dispose();
            }
            if (_slotStatus.TryGetValue(size, out Label lbl))
            {
                lbl.Text = "(empty)";
                lbl.ForeColor = SystemColors.GrayText;
            }
        }

        private void OnExportClicked()
        {
            if (_slotBitmaps.Count == 0)
            {
                MessageBox.Show("Assign at least one slot before exporting.",
                    "Icon Editor", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Build the same fully-resolved per-size dictionary that WriteIcoFromSizedSources
            // will write, so the preview shows the exact pixels going into the .ico (including
            // any sizes auto-rendered from the largest filled slot).
            IDictionary<int, Bitmap> resolved = null;
            try
            {
                resolved = IconWriter.ResolveSources(_slotBitmaps);

                using (var preview = new IconPreviewForm(resolved))
                {
                    if (preview.ShowDialog(this) != DialogResult.OK) return;
                }

                string outputPath;
                using (var dlg = new SaveFileDialog())
                {
                    dlg.Title = "Export Icon As";
                    dlg.Filter = "Icon files (*.ico)|*.ico";
                    dlg.DefaultExt = "ico";
                    dlg.AddExtension = true;
                    dlg.OverwritePrompt = true;
                    dlg.FileName = "icon.ico";
                    if (dlg.ShowDialog() != DialogResult.OK) return;
                    outputPath = dlg.FileName;
                }

                IconWriter.WriteIcoFromSizedSources(resolved, outputPath);
                MessageBox.Show("Icon saved:\n" + outputPath, "Icon Editor",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save icon:\n" + ex.Message, "Icon Editor",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (resolved != null)
                {
                    foreach (var b in resolved.Values) b?.Dispose();
                }
            }
        }

        // --- Undo plumbing ---

        private void PushUndo(IEnumerable<int> sizes)
        {
            var entry = new Dictionary<int, SlotSnapshot>();
            foreach (int s in sizes)
            {
                if (!entry.ContainsKey(s)) entry[s] = CaptureSlot(s);
            }
            _undoStack.Push(entry);
            // Cap stack depth so an extremely long session can't grow memory unboundedly.
            // The bottom (oldest) entry is the one to drop, but Stack<T> doesn't expose that
            // efficiently, so rebuild as an array if we exceed the cap (rare path).
            if (_undoStack.Count > MaxUndoDepth)
            {
                var arr = _undoStack.ToArray(); // top first
                // Discard the oldest entry and dispose its captured bitmaps.
                var dropped = arr[arr.Length - 1];
                foreach (var snap in dropped.Values) snap.Bitmap?.Dispose();
                _undoStack.Clear();
                for (int i = arr.Length - 2; i >= 0; i--) _undoStack.Push(arr[i]);
            }
        }

        private SlotSnapshot CaptureSlot(int size)
        {
            var snap = new SlotSnapshot();
            if (_slotBitmaps.TryGetValue(size, out Bitmap bmp) && bmp != null)
            {
                snap.Bitmap = (Bitmap)bmp.Clone();
            }
            if (_slotStatus.TryGetValue(size, out Label lbl))
            {
                snap.StatusText = lbl.Text;
                snap.StatusColor = lbl.ForeColor;
            }
            else
            {
                snap.StatusText = "(empty)";
                snap.StatusColor = SystemColors.GrayText;
            }
            return snap;
        }

        private void Undo()
        {
            if (_undoStack.Count == 0) return;
            var entry = _undoStack.Pop();
            foreach (var kv in entry)
            {
                int size = kv.Key;
                var snap = kv.Value;
                if (snap.Bitmap == null)
                {
                    // Slot was empty before the action — clear it back out (without pushing
                    // a new undo entry, since we're undoing).
                    ClearSlotInternal(size);
                }
                else
                {
                    // Restore the cloned bitmap and the recorded status. SetSlotBitmap takes
                    // ownership, so we null out the snapshot reference to avoid double-dispose.
                    SetSlotBitmap(size, snap.Bitmap, snap.StatusText);
                    if (_slotStatus.TryGetValue(size, out Label lbl)) lbl.ForeColor = snap.StatusColor;
                    snap.Bitmap = null;
                }
            }
        }

        // Catches Ctrl+Z anywhere inside the pane (including immediately after one of our
        // buttons fires, since the button keeps focus).
        //
        // PowerPoint doesn't expose a way to push our pane's actions into its own undo stack,
        // so the two stacks are chained: we pop ours first, and once it's empty we forward
        // Ctrl+Z to PowerPoint's built-in undo via CommandBars.ExecuteMso("Undo"). This means a
        // user holding Ctrl+Z walks back through their pane actions and then keeps walking
        // back through their slide actions, which is the closest we can get to a single
        // unified undo experience.
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.Z))
            {
                if (_undoStack.Count > 0)
                {
                    Undo();
                    return true;
                }
                if (TryForwardUndoToPowerPoint())
                {
                    return true;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private static bool TryForwardUndoToPowerPoint()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                var bars = app?.CommandBars;
                if (bars == null) return false;
                bars.ExecuteMso("Undo");
                return true;
            }
            catch
            {
                return false;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                foreach (var bmp in _slotBitmaps.Values) bmp?.Dispose();
                _slotBitmaps.Clear();
                foreach (var pb in _slotPreviews.Values) pb.Image?.Dispose();
                while (_undoStack.Count > 0)
                {
                    var entry = _undoStack.Pop();
                    foreach (var snap in entry.Values) snap.Bitmap?.Dispose();
                }
            }
            base.Dispose(disposing);
        }
    }

    // A single icon-size row in the editor pane. Performs its own responsive layout based on
    // its current Width:
    //   - Wide:   [preview thumbnail] [size + status text] [Assign / Clear buttons]
    //   - Narrow: [preview thumbnail]                      [Assign / Clear buttons]
    //             [size + status text spanning beneath the thumbnail, wrapping as needed]
    // A MinimumSize is also published so that when the host pane is shrunk below the point
    // where even the narrow layout fits, the host's AutoScroll engages a horizontal scrollbar
    // rather than the row clipping/overlapping its controls.
    internal sealed class SlotRowControl : Panel
    {
        private const int Gap = 8;
        // When the inline-text column would be narrower than this, switch to the stacked
        // narrow layout where the text wraps under the preview thumbnail instead.
        private const int MinInlineTextWidth = 70;

        private readonly int _previewSize;
        public PictureBox Preview { get; }
        public Label SizeLabel { get; }
        public Label StatusLabel { get; }
        public Button AssignButton { get; }
        public Button ClearButton { get; }
        private readonly FlowLayoutPanel _buttons;

        public SlotRowControl(int size, Font baseFont, int previewSize)
        {
            _previewSize = previewSize;
            // Bottom Padding adds visual spacing between this row and the next docked
            // sibling (Dock=Top ignores Margin). Children are positioned relative to (0,0)
            // so they sit above this padding.
            Padding = new Padding(0, 0, 0, 12);

            Preview = new PictureBox
            {
                Width = previewSize,
                Height = previewSize,
                BorderStyle = BorderStyle.FixedSingle,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackgroundImage = GetCheckerTile(),
                BackgroundImageLayout = ImageLayout.Tile,
            };
            SizeLabel = new Label
            {
                Text = size + " \u00D7 " + size,
                Font = new Font(baseFont, FontStyle.Bold),
                AutoSize = true,
            };
            StatusLabel = new Label
            {
                Text = "(empty)",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
            };
            AssignButton = new RoundedAccentButton(Color.FromArgb(0, 120, 215))
            {
                Text = "Assign",
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 2),
            };
            ClearButton = new Button
            {
                Text = "Clear",
                AutoSize = true,
                Margin = new Padding(0),
                UseVisualStyleBackColor = true,
            };
            _buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
            };
            _buttons.Controls.Add(AssignButton);
            _buttons.Controls.Add(ClearButton);

            Controls.Add(Preview);
            Controls.Add(SizeLabel);
            Controls.Add(StatusLabel);
            Controls.Add(_buttons);

            // When the status text changes (e.g. "(empty)" → "Assigned: Slide 1, shape rect 24")
            // the row may need a different height (taller in narrow mode if it wraps to more
            // lines). Force a re-measure.
            StatusLabel.TextChanged += (s, e) =>
            {
                int desiredH = ComputeContentHeight(Width) + Padding.Vertical;
                if (Height != desiredH) Height = desiredH;
                else DoResponsiveLayout();
            };
        }

        // Self-sizing height-for-width: when Dock=Top stretches our Width to fit the parent,
        // OnSizeChanged fires and we set our own Height to match the layout's needs (inline
        // vs narrow mode). The framework natively re-docks subsequent siblings.
        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            int desiredH = ComputeContentHeight(Width) + Padding.Vertical;
            if (Height != desiredH)
            {
                Height = desiredH;
                // Setting Height re-fires OnSizeChanged; the if-Height-changed guard above
                // breaks the loop on the next pass.
                return;
            }
            DoResponsiveLayout();
        }

        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            // Force a re-measure on the next OnSizeChanged.
            if (Width > 0)
            {
                int desiredH = ComputeContentHeight(Width) + Padding.Vertical;
                if (Height != desiredH) Height = desiredH;
                else DoResponsiveLayout();
            }
        }

        // Pure measurement: how tall does our content (excluding Padding) need to be
        // at this width? Doesn't mutate any child controls.
        private int ComputeContentHeight(int width)
        {
            int w = Math.Max(width, _previewSize + Gap + 1);
            Size btnsPref = _buttons.GetPreferredSize(Size.Empty);
            int btnW = btnsPref.Width;
            int btnH = btnsPref.Height;

            int inlineTextX = _previewSize + Gap;
            int inlineTextMaxW = w - inlineTextX - Gap - btnW;
            bool inline = inlineTextMaxW >= MinInlineTextWidth;

            int textMaxW = inline ? inlineTextMaxW : Math.Max(40, w);
            int sizeH = MeasureLabelHeight(SizeLabel.Text, SizeLabel.Font, textMaxW);
            int statusH = MeasureLabelHeight(StatusLabel.Text, StatusLabel.Font, textMaxW);
            int interLineGap = inline ? 2 : -3;
            int textBlockH = sizeH + interLineGap + statusH;

            return inline
                ? Math.Max(_previewSize, Math.Max(textBlockH + 4, btnH))
                : Math.Max(_previewSize + Gap + textBlockH, btnH);
        }

        private static int MeasureLabelHeight(string text, Font font, int maxWidth)
        {
            int w = Math.Max(1, maxWidth);
            Size measured = TextRenderer.MeasureText(text ?? string.Empty, font,
                new Size(w, int.MaxValue),
                TextFormatFlags.WordBreak | TextFormatFlags.NoPadding);
            return measured.Height + 3;
        }

        private void DoResponsiveLayout()
        {
            int width = Math.Max(Width, MinimumSize.Width);
            Size btnsPref = _buttons.GetPreferredSize(Size.Empty);
            int btnW = btnsPref.Width;
            int btnH = btnsPref.Height;

            // Buttons hug the right edge of the row, vertically aligned with the top of the
            // preview thumbnail in both layout modes.
            _buttons.Size = btnsPref;
            int buttonsX = Math.Max(_previewSize + Gap, width - btnW);
            _buttons.Location = new Point(buttonsX, 0);

            Preview.Location = new Point(0, 0);

            int inlineTextX = _previewSize + Gap;
            int inlineTextMaxW = width - inlineTextX - Gap - btnW;
            bool inline = inlineTextMaxW >= MinInlineTextWidth;

            int textX, textMaxW, textY;
            if (inline)
            {
                textX = inlineTextX;
                textMaxW = inlineTextMaxW;
                textY = 4;
            }
            else
            {
                textX = 0;
                textMaxW = Math.Max(40, width);
                textY = _previewSize + Gap;
            }

            SizeLabel.MaximumSize = new Size(textMaxW, 0);
            SizeLabel.Location = new Point(textX, textY);
            StatusLabel.MaximumSize = new Size(textMaxW, 0);
            int interLineGap = inline ? 2 : -3;
            StatusLabel.Location = new Point(textX, SizeLabel.Bottom + interLineGap);

            // MinimumSize so the host's AutoScroll engages horizontally at the right point.
            int minWidth = _previewSize + Gap + btnW;
            if (MinimumSize.Width != minWidth)
            {
                MinimumSize = new Size(minWidth, 0);
            }

            // Push our minimum width up to the host UserControl. Dock=Top children always fit
            // horizontally to the parent's client width — they don't overflow on their own —
            // so the host needs an explicit AutoScrollMinSize.Width for the horizontal
            // scrollbar to appear when the pane is narrower than this row's minimum.
            for (Control p = Parent; p != null; p = p.Parent)
            {
                if (p is IconEditorPaneControl pane)
                {
                    pane.NotifySlotRowMinWidth(minWidth);
                    break;
                }
            }
        }

        private static Bitmap s_checker;
        private static Bitmap GetCheckerTile()
        {
            if (s_checker != null) return s_checker;
            const int tile = 16;
            var bmp = new Bitmap(tile * 2, tile * 2);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                using (var brush = new SolidBrush(Color.FromArgb(220, 220, 220)))
                {
                    g.FillRectangle(brush, 0, 0, tile, tile);
                    g.FillRectangle(brush, tile, tile, tile, tile);
                }
            }
            s_checker = bmp;
            return s_checker;
        }
    }

    // Button subclass that mimics the Windows 11 themed button shape (rounded ~4px corners,
    // standard horizontal padding) but with a fully customizable background color and white
    // foreground. Standard WinForms buttons can't be tinted while keeping their themed look —
    // setting BackColor on a non-flat button is silently ignored on themed Windows, and
    // FlatStyle=Flat strips the rounding/padding. This owner-drawn variant restores both.
    internal sealed class RoundedAccentButton : Button
    {
        private const int CornerRadius = 4;
        private readonly Color _baseColor;
        private bool _hover;
        private bool _pressed;

        public RoundedAccentButton(Color baseColor)
        {
            _baseColor = baseColor;
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            FlatAppearance.MouseOverBackColor = Color.Transparent;
            FlatAppearance.MouseDownBackColor = Color.Transparent;
            BackColor = baseColor;
            ForeColor = Color.White;
            SetStyle(ControlStyles.UserPaint
                | ControlStyles.AllPaintingInWmPaint
                | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw, true);
        }

        // Match the size a standard themed Button would auto-size to with the same text and
        // font, so a row of mixed colored/standard buttons line up exactly.
        public override Size GetPreferredSize(Size proposedSize)
        {
            using (var probe = new Button { Text = this.Text, Font = this.Font, AutoSize = true })
            {
                return probe.GetPreferredSize(proposedSize);
            }
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; _pressed = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs mevent) { _pressed = true; Invalidate(); base.OnMouseDown(mevent); }
        protected override void OnMouseUp(MouseEventArgs mevent) { _pressed = false; Invalidate(); base.OnMouseUp(mevent); }

        protected override void OnPaint(PaintEventArgs pevent)
        {
            var g = pevent.Graphics;
            // Parent fills the button's rectangular bounding box first so the four rounded
            // corners show through to the parent's background color.
            if (Parent != null)
            {
                using (var bg = new SolidBrush(Parent.BackColor)) g.FillRectangle(bg, ClientRectangle);
            }

            Color fill = !Enabled ? ControlPaint.Light(_baseColor, 0.4f)
                       : _pressed ? ControlPaint.Dark(_baseColor, 0.08f)
                       : _hover   ? ControlPaint.Light(_baseColor, 0.10f)
                       : _baseColor;

            var rect = new Rectangle(0, 0, Width - 1, Height - 1);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = BuildRoundedRect(rect, CornerRadius))
            using (var brush = new SolidBrush(fill))
            {
                g.FillPath(brush, path);
            }
            g.SmoothingMode = SmoothingMode.Default;

            TextRenderer.DrawText(g, Text, Font, ClientRectangle, ForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter
                | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
        }

        private static GraphicsPath BuildRoundedRect(Rectangle r, int radius)
        {
            int d = radius * 2;
            var path = new GraphicsPath();
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
