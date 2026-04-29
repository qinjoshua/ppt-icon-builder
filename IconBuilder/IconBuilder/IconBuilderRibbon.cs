using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace IconBuilder
{
    [ComVisible(true)]
    public class IconBuilderRibbon : Office.IRibbonExtensibility
    {
        private const string MenuLabel = "Save as Icon (.ico)...";
        private const string ButtonId = "IconBuilder_SaveAsIco";
        private const string CallbackName = "OnSaveAsIco";

        // Icon sizes the user can insert as guide squares. Selection is tracked here and persists
        // for the lifetime of the add-in instance (defaults to all sizes selected).
        private static readonly int[] GuideSizes = { 16, 20, 24, 32, 40, 48, 64, 256 };
        private readonly System.Collections.Generic.Dictionary<int, bool> _guideSizeSelected;
        private Office.IRibbonUI _ribbon;

        // Custom responsive collapse for the Guide Sizes group: when the active PowerPoint
        // window's width drops below CollapseThresholdPoints, the inline checkbox grid is hidden
        // and replaced by a single dropdown menu button. Office's built-in group auto-collapse
        // is unreliable for groups composed of many small controls, so we drive it ourselves
        // via a low-frequency polling timer (PowerPoint's interop has no direct WindowResize
        // event we can hook).
        private const float CollapseThresholdPoints = 500f;
        private const float ExpandHysteresisPoints = 40f;
        private System.Windows.Forms.Timer _responsiveTimer;
        private bool _guideSizesCollapsed;

        public IconBuilderRibbon()
        {
            _guideSizeSelected = new System.Collections.Generic.Dictionary<int, bool>();
            foreach (int s in GuideSizes) _guideSizeSelected[s] = true;
        }

        public string GetCustomUI(string RibbonID)
        {
            // The customUI XML lives in IconBuilderRibbon.xml (embedded resource) so it can be
            // edited without rebuilding string concatenations and so the structural definition
            // is in a place where XML tooling can lint/validate it. The single piece that
            // genuinely is dynamic — the per-size <checkBox> elements — is generated from
            // GuideSizes here so that the supported size set has exactly one source of truth.
            string template = LoadRibbonTemplate();

            var inlineSb = new System.Text.StringBuilder();
            int half = (GuideSizes.Length + 1) / 2;
            inlineSb.Append("            <box id=\"GuideSizesCol1\" boxStyle=\"vertical\">\r\n");
            for (int i = 0; i < half; i++)
            {
                int s = GuideSizes[i];
                inlineSb.AppendFormat(
                    "              <checkBox id=\"GuideChk_{0}\" label=\"{0} \u00D7 {0}\" onAction=\"OnGuideSizeToggle\" getPressed=\"GetGuideSizePressed\"/>\r\n",
                    s);
            }
            inlineSb.Append("            </box>\r\n");
            inlineSb.Append("            <box id=\"GuideSizesCol2\" boxStyle=\"vertical\">\r\n");
            for (int i = half; i < GuideSizes.Length; i++)
            {
                int s = GuideSizes[i];
                inlineSb.AppendFormat(
                    "              <checkBox id=\"GuideChk_{0}\" label=\"{0} \u00D7 {0}\" onAction=\"OnGuideSizeToggle\" getPressed=\"GetGuideSizePressed\"/>\r\n",
                    s);
            }
            inlineSb.Append("            </box>");

            var menuSb = new System.Text.StringBuilder();
            menuSb.Append("          <menu id=\"GuideSizesCollapsed\" size=\"large\" label=\"Guide Sizes\" imageMso=\"GroupShapes\" itemSize=\"normal\" getVisible=\"GetGuideSizesCollapsedVisible\">\r\n");
            for (int i = 0; i < GuideSizes.Length; i++)
            {
                int s = GuideSizes[i];
                menuSb.AppendFormat(
                    "            <checkBox id=\"GuideMenuChk_{0}\" label=\"{0} \u00D7 {0}\" onAction=\"OnGuideSizeToggle\" getPressed=\"GetGuideSizePressed\"/>\r\n",
                    s);
            }
            menuSb.Append("          </menu>");

            return template
                .Replace("{GuideSizesInlineXml}", inlineSb.ToString())
                .Replace("{GuideSizesMenuXml}", menuSb.ToString())
                .Replace("{CallbackName}", CallbackName)
                .Replace("{ButtonId}", ButtonId)
                .Replace("{MenuLabel}", MenuLabel);
        }

        // Reads the customUI XML template that is embedded next to this class. Robust against
        // assembly-name / namespace drift and against the EmbeddedResource entry being lost
        // from the .csproj (e.g. when Visual Studio rewrites the legacy project file): falls
        // back first to suffix-matching any embedded resource ending in "IconBuilderRibbon.xml",
        // then to reading the file from disk next to the running assembly.
        private static string LoadRibbonTemplate()
        {
            var asm = typeof(IconBuilderRibbon).Assembly;
            const string PreferredName = "IconBuilder.IconBuilderRibbon.xml";
            const string FileName = "IconBuilderRibbon.xml";

            string[] resourceNames = asm.GetManifestResourceNames();

            string match = null;
            foreach (var n in resourceNames)
            {
                if (string.Equals(n, PreferredName, StringComparison.Ordinal)) { match = n; break; }
            }
            if (match == null)
            {
                foreach (var n in resourceNames)
                {
                    if (n.EndsWith(FileName, StringComparison.OrdinalIgnoreCase)) { match = n; break; }
                }
            }

            if (match != null)
            {
                using (var stream = asm.GetManifestResourceStream(match))
                using (var reader = new StreamReader(stream, System.Text.Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }

            // Fallback: look for the file on disk next to the loaded assembly. Useful during
            // development when the EmbeddedResource entry has been stripped from the project.
            try
            {
                string asmPath = new Uri(asm.CodeBase).LocalPath;
                string asmDir = Path.GetDirectoryName(asmPath);
                if (!string.IsNullOrEmpty(asmDir))
                {
                    string[] candidates =
                    {
                        Path.Combine(asmDir, FileName),
                        Path.Combine(asmDir, "..", FileName),
                        Path.Combine(asmDir, "..", "..", FileName),
                    };
                    foreach (var p in candidates)
                    {
                        if (File.Exists(p)) return File.ReadAllText(p, System.Text.Encoding.UTF8);
                    }
                }
            }
            catch { /* fall through to throw */ }

            throw new InvalidOperationException(
                "Embedded ribbon template not found: " + PreferredName
                + ". Available resources: " + string.Join(", ", resourceNames));
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            // If we get here, Office accepted our customUI XML. Otherwise the XML was rejected
            // (typically because of an unknown idMso/imageMso) and this callback never fires.
            try
            {
                string logPath = Path.Combine(Path.GetTempPath(), "IconBuilder_RibbonLoad.log");
                File.AppendAllText(logPath,
                    DateTime.Now.ToString("o") + " Ribbon loaded successfully\r\n");
            }
            catch { }

            // When the user closes the Icon Editor pane via its X button, the toggle button on
            // the ribbon needs to refresh its pressed state.
            try
            {
                var addin = Globals.ThisAddIn;
                if (addin != null)
                {
                    addin.EditorPaneVisibilityChanged += (s, e) =>
                    {
                        try { _ribbon?.InvalidateControl("IconBuilder_PaneBtn"); } catch { }
                    };
                }
            }
            catch { }

            StartResponsiveTimer();
        }

        // Starts a low-frequency polling timer that drives the Guide Sizes group's custom
        // responsive collapse. PPT interop has no WindowResize event, so polling is the
        // pragmatic option; 4 Hz is fast enough that the user sees the swap as immediate while
        // being light enough to be invisible to the CPU.
        private void StartResponsiveTimer()
        {
            try
            {
                if (_responsiveTimer != null) return;
                _responsiveTimer = new System.Windows.Forms.Timer { Interval = 250 };
                _responsiveTimer.Tick += (s, e) => UpdateResponsiveState();
                _responsiveTimer.Start();
                // Run once immediately so the initial state is correct without waiting 250ms.
                UpdateResponsiveState();
            }
            catch { }
        }

        // Reads the active PowerPoint window's width (in points; 1pt = 1/72 inch) and toggles
        // the collapsed/expanded representation when the threshold is crossed. A small
        // hysteresis band prevents rapid flip-flopping when the user drags the window edge
        // exactly across the threshold.
        private void UpdateResponsiveState()
        {
            try
            {
                var addin = Globals.ThisAddIn;
                var app = addin?.Application;
                if (app == null) return;

                float width;
                try { width = (float)app.ActiveWindow.Width; }
                catch { return; } // No active presentation window yet.

                bool nowCollapsed = _guideSizesCollapsed
                    ? width < CollapseThresholdPoints + ExpandHysteresisPoints
                    : width < CollapseThresholdPoints;

                if (nowCollapsed != _guideSizesCollapsed)
                {
                    _guideSizesCollapsed = nowCollapsed;
                    try { _ribbon?.InvalidateControl("GuideSizesExpanded"); } catch { }
                    try { _ribbon?.InvalidateControl("GuideSizesCollapsed"); } catch { }
                }
            }
            catch { }
        }

        public bool GetGuideSizesExpandedVisible(Office.IRibbonControl control)
        {
            return !_guideSizesCollapsed;
        }

        public bool GetGuideSizesCollapsedVisible(Office.IRibbonControl control)
        {
            return _guideSizesCollapsed;
        }

        // --- Icon Editor pane toggle callbacks ---

        public bool GetEditorPanePressed(Office.IRibbonControl control)
        {
            try { return Globals.ThisAddIn?.EditorTaskPane?.Visible ?? false; }
            catch { return false; }
        }

        public void OnToggleEditorPane(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                var pane = Globals.ThisAddIn?.EditorTaskPane;
                if (pane != null) pane.Visible = pressed;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to toggle Icon Editor pane:\n" + ex.Message,
                    "Icon Editor", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // --- Guide-size dropdown checkbox callbacks ---

        public bool GetGuideSizePressed(Office.IRibbonControl control)
        {
            int size = ParseGuideSizeFromId(control?.Id);
            return size > 0 && _guideSizeSelected.TryGetValue(size, out bool v) && v;
        }

        public void OnGuideSizeToggle(Office.IRibbonControl control, bool pressed)
        {
            int size = ParseGuideSizeFromId(control?.Id);
            if (size > 0) _guideSizeSelected[size] = pressed;
        }

        private static int ParseGuideSizeFromId(string id)
        {
            if (string.IsNullOrEmpty(id)) return 0;
            int idx = id.LastIndexOf('_');
            if (idx < 0 || idx >= id.Length - 1) return 0;
            return int.TryParse(id.Substring(idx + 1), out int n) ? n : 0;
        }

        // Inserts square guides on the active slide sized to icon pixel dimensions, using whichever
        // sizes the user has currently checked in the dropdown menu (defaults to all sizes).
        // 96 DPI conversion is used (1 pixel = 0.75 points). Each square is unfilled with a colored
        // outline and a label below describing its pixel size.
        public void OnInsertGuideSquares(Office.IRibbonControl control)
        {
            const float PxToPt = 72f / 96f; // 96 DPI => 0.75 pt per pixel.

            // Determine selected sizes in descending order (largest leftmost so labels read naturally).
            var selected = new List<int>();
            foreach (int s in GuideSizes)
            {
                if (_guideSizeSelected.TryGetValue(s, out bool on) && on) selected.Add(s);
            }
            selected.Sort((a, b) => b.CompareTo(a));

            if (selected.Count == 0)
            {
                MessageBox.Show("Please select at least one size from the dropdown.",
                    "Insert Guide Squares", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Color palette indexed by size so each size gets a stable distinguishing color.
            var colorBySize = new Dictionary<int, int>
            {
                { 256, ColorRgb(0x1F, 0x77, 0xB4) }, // blue
                {  64, ColorRgb(0x94, 0x67, 0xBD) }, // purple
                {  48, ColorRgb(0xFF, 0x7F, 0x0E) }, // orange
                {  40, ColorRgb(0x8C, 0x56, 0x4B) }, // brown
                {  32, ColorRgb(0x2C, 0xA0, 0x2C) }, // green
                {  24, ColorRgb(0xE3, 0x77, 0xC2) }, // pink
                {  20, ColorRgb(0x17, 0xBE, 0xCF) }, // cyan
                {  16, ColorRgb(0xD6, 0x27, 0x28) }, // red
            };

            try
            {
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Slide slide = null;
                try { slide = (PowerPoint.Slide)app.ActiveWindow.View.Slide; } catch { }
                if (slide == null)
                {
                    MessageBox.Show("Open a slide first.", "Insert Guide Squares",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var pres = app.ActivePresentation;
                float slideW = pres.PageSetup.SlideWidth;
                float slideH = pres.PageSetup.SlideHeight;

                const float Margin = 18f;     // 0.25 inch in points
                const float Gap = 12f;        // gap between squares
                const float LabelHeight = 14f;

                // Total width occupied by guides: largest first then descending, with gaps.
                float totalW = 0f;
                for (int i = 0; i < selected.Count; i++)
                {
                    totalW += selected[i] * PxToPt;
                    if (i > 0) totalW += Gap;
                }

                float startX = Math.Max(Margin, (slideW - totalW) / 2f);
                float maxSquarePt = selected[0] * PxToPt;
                float topY = Math.Max(Margin, (slideH - maxSquarePt - LabelHeight - 4f) / 2f);

                var shapes = slide.Shapes;
                var newGuideShapes = new List<PowerPoint.Shape>();

                float cursorX = startX;
                for (int i = 0; i < selected.Count; i++)
                {
                    int px = selected[i];
                    float sidePt = px * PxToPt;
                    // Bottom-align all squares so their baselines match.
                    float y = topY + (maxSquarePt - sidePt);

                    int color = colorBySize.TryGetValue(px, out int c) ? c : ColorRgb(0x55, 0x55, 0x55);

                    var rect = shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        cursorX, y, sidePt, sidePt);
                    rect.Name = $"IconBuilderGuide_{px}px";
                    rect.Fill.Visible = Office.MsoTriState.msoFalse;
                    rect.Line.Visible = Office.MsoTriState.msoTrue;
                    rect.Line.Weight = 1.0f;
                    rect.Line.ForeColor.RGB = color;

                    var label = shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        cursorX, y + sidePt + 2f, sidePt, LabelHeight);
                    label.Name = $"IconBuilderGuideLabel_{px}px";
                    label.Line.Visible = Office.MsoTriState.msoFalse;
                    label.Fill.Visible = Office.MsoTriState.msoFalse;
                    var tf = label.TextFrame;
                    tf.MarginLeft = 0; tf.MarginRight = 0; tf.MarginTop = 0; tf.MarginBottom = 0;
                    tf.WordWrap = Office.MsoTriState.msoFalse;
                    var tr = tf.TextRange;
                    tr.Text = $"{px}\u00D7{px}";
                    tr.Font.Size = 9;
                    tr.Font.Bold = Office.MsoTriState.msoTrue;
                    tr.Font.Color.RGB = color;
                    tr.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;

                    newGuideShapes.Add(rect);
                    newGuideShapes.Add(label);

                    cursorX += sidePt + Gap;
                }

                // Select the new guides so the user can immediately move/group them as a unit.
                if (newGuideShapes.Count > 0)
                {
                    try
                    {
                        var names = new string[newGuideShapes.Count];
                        for (int i = 0; i < newGuideShapes.Count; i++) names[i] = newGuideShapes[i].Name;
                        slide.Shapes.Range(names).Select();
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to insert guide squares:\n" + ex.Message,
                    "Insert Guide Squares", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static int ColorRgb(int r, int g, int b)
        {
            // PowerPoint uses BGR in the RGB property.
            return (b << 16) | (g << 8) | r;
        }

        public void OnSaveAsIco(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = null;
                try { sel = app.ActiveWindow?.Selection; } catch { sel = null; }

                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please right-click on a shape, picture, or object first.",
                        "Save as Icon", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (range == null || range.Count == 0)
                {
                    MessageBox.Show("No shape selected.", "Save as Icon",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string defaultName = SafeFileName(range.Count == 1 ? range[1].Name : "icon");

                string tempPng = Path.Combine(Path.GetTempPath(),
                    "IconBuilder_" + Guid.NewGuid().ToString("N") + ".png");

                // Render per-size bitmaps once up front so the preview window can show what
                // the user is about to save, pixel-for-pixel. Same rendering pipeline that
                // WriteIcoFromImage uses internally (trim transparent borders, then bicubic
                // letterbox into each target size), so the preview matches the saved file.
                var previewBitmaps = new System.Collections.Generic.Dictionary<int, Bitmap>();
                try
                {
                    ExportShapeRangeToPng(range, tempPng);

                    using (var src = Image.FromFile(tempPng))
                    using (var trimmed = IconWriter.TrimTransparentBorder(src))
                    {
                        foreach (int size in IconWriter.BmpSizes)
                            previewBitmaps[size] = IconWriter.ResizeTo32bpp(trimmed, size);
                        previewBitmaps[IconWriter.PngSize] =
                            IconWriter.ResizeTo32bpp(trimmed, IconWriter.PngSize);
                    }

                    using (var preview = new IconPreviewForm(previewBitmaps))
                    {
                        if (preview.ShowDialog() != DialogResult.OK) return;
                    }

                    string outputPath;
                    using (var dlg = new SaveFileDialog())
                    {
                        dlg.Title = "Save Icon As";
                        dlg.Filter = "Icon files (*.ico)|*.ico";
                        dlg.DefaultExt = "ico";
                        dlg.AddExtension = true;
                        dlg.OverwritePrompt = true;
                        dlg.FileName = defaultName + ".ico";
                        if (dlg.ShowDialog() != DialogResult.OK) return;
                        outputPath = dlg.FileName;
                    }

                    IconWriter.WriteIcoFromSizedSources(previewBitmaps, outputPath);

                    MessageBox.Show("Icon saved:\n" + outputPath, "Save as Icon",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    foreach (var b in previewBitmaps.Values) b?.Dispose();
                    try { if (File.Exists(tempPng)) File.Delete(tempPng); } catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save icon:\n" + ex.Message,
                    "Save as Icon", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Exports a ShapeRange to PNG. For multi-shape selections (not already a single group),
        // temporarily group the shapes so the export captures all of them as one image, then ungroup.
        //
        // Strategy: use ppRelativeToSlide and pass ScaleWidth/ScaleHeight that match the SLIDE's
        // aspect ratio, scaled so the shape's longest side comes out near a target pixel size.
        // This is the only mode whose proportion behavior is reliable across Office builds:
        // ppScaleXY does not always treat its parameters as independent pixel dimensions, which
        // causes horizontal/vertical compression of shapes on non-square slides.
        internal static void ExportShapeRangeToPng(PowerPoint.ShapeRange range, string pngPath)
        {
            const int targetShapePx = 512;
            const PowerPoint.PpShapeFormat fmt = PowerPoint.PpShapeFormat.ppShapeFormatPNG;
            const PowerPoint.PpExportMode mode = PowerPoint.PpExportMode.ppRelativeToSlide;

            var app = Globals.ThisAddIn.Application;
            var pres = app.ActivePresentation;
            float slideW = pres.PageSetup.SlideWidth;
            float slideH = pres.PageSetup.SlideHeight;

            if (range.Count <= 1)
            {
                var s = range[1];
                var (sw, sh) = ComputeSlideRenderSize(slideW, slideH, s.Width, s.Height, targetShapePx);
                s.Export(pngPath, fmt, sw, sh, mode);
                return;
            }

            string[] originalNames = new string[range.Count];
            for (int i = 1; i <= range.Count; i++) originalNames[i - 1] = range[i].Name;

            PowerPoint.Shape grouped = null;
            bool weGrouped = false;
            try
            {
                try
                {
                    grouped = range.Group();
                    weGrouped = true;
                }
                catch
                {
                    var s = range[1];
                    var (sw0, sh0) = ComputeSlideRenderSize(slideW, slideH, s.Width, s.Height, targetShapePx);
                    range.Export(pngPath, fmt, sw0, sh0, mode);
                    return;
                }

                var (sw, sh) = ComputeSlideRenderSize(slideW, slideH, grouped.Width, grouped.Height, targetShapePx);
                grouped.Export(pngPath, fmt, sw, sh, mode);
            }
            finally
            {
                if (weGrouped && grouped != null)
                {
                    try { grouped.Ungroup(); } catch { }

                    try
                    {
                        var slide = (PowerPoint.Slide)Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                        var restored = slide.Shapes.Range(originalNames);
                        restored.Select();
                    }
                    catch { }
                }
            }
        }

        // Returns the pixel dimensions to pass to Shape.Export in ppRelativeToSlide mode such that:
        //   1. The slide's aspect ratio is preserved in the rendering (no shape distortion).
        //   2. The shape's longest side, after cropping out of that rendered slide, is ~targetShapePx.
        private static (int slideWidthPx, int slideHeightPx) ComputeSlideRenderSize(
            float slideW, float slideH, float shapeW, float shapeH, int targetShapePx)
        {
            if (slideW <= 0 || slideH <= 0 || shapeW <= 0 || shapeH <= 0)
                return (targetShapePx, targetShapePx);

            // Pixels-per-point such that shape's longest side equals targetShapePx.
            double pxPerPt = targetShapePx / Math.Max(shapeW, shapeH);
            int slidePxW = Math.Max(1, (int)Math.Round(slideW * pxPerPt));
            int slidePxH = Math.Max(1, (int)Math.Round(slideH * pxPerPt));
            return (slidePxW, slidePxH);
        }

        private static (int width, int height) ComputeAspectPreservingSize(float shapeW, float shapeH, int maxDim)
        {
            if (shapeW <= 0 || shapeH <= 0) return (maxDim, maxDim);
            double scale = maxDim / Math.Max(shapeW, shapeH);
            int w = Math.Max(1, (int)Math.Round(shapeW * scale));
            int h = Math.Max(1, (int)Math.Round(shapeH * scale));
            return (w, h);
        }

        private static string SafeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "icon";
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }
    }
}
