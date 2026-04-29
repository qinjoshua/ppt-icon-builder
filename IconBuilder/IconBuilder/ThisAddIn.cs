using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace IconBuilder
{
    public partial class ThisAddIn
    {
        // PowerPoint creates one DocumentWindow per visible presentation window. Each window
        // needs its own CustomTaskPane (otherwise the pane only appears in the window where
        // CustomTaskPanes.Add was called from). We lazily create a pane the first time a
        // window becomes active and dispose it when the window is closed.
        private sealed class PaneEntry
        {
            public Microsoft.Office.Tools.CustomTaskPane TaskPane;
            public IconEditorPaneControl Pane;
        }

        private readonly Dictionary<PowerPoint.DocumentWindow, PaneEntry> _panes
            = new Dictionary<PowerPoint.DocumentWindow, PaneEntry>();

        public event EventHandler EditorPaneVisibilityChanged;

        // Convenience accessors that resolve to the *active* window's pane. Returns null when
        // there's no active window or the pane couldn't be created.
        public Microsoft.Office.Tools.CustomTaskPane EditorTaskPane
        {
            get { return GetOrCreateActivePaneEntry()?.TaskPane; }
        }

        public IconEditorPaneControl EditorPane
        {
            get { return GetOrCreateActivePaneEntry()?.Pane; }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowActivate += OnWindowActivate;
            // Create a pane for the window that's already active at startup, if any.
            try
            {
                if (this.Application.Windows.Count > 0)
                {
                    EnsurePaneForWindow(this.Application.ActiveWindow);
                }
            }
            catch
            {
                // Application.ActiveWindow throws if no window exists yet — that's fine,
                // OnWindowActivate will create one when a window becomes active.
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try { this.Application.WindowActivate -= OnWindowActivate; } catch { }
            foreach (var entry in _panes.Values)
            {
                try { entry.TaskPane.Dispose(); } catch { }
            }
            _panes.Clear();
        }

        private void OnWindowActivate(PowerPoint.Presentation pres, PowerPoint.DocumentWindow wn)
        {
            try { EnsurePaneForWindow(wn); }
            catch { }
            // Take this opportunity to GC panes whose windows have been closed.
            try { CleanupClosedWindows(); } catch { }
        }

        private PaneEntry EnsurePaneForWindow(PowerPoint.DocumentWindow wn)
        {
            if (wn == null) return null;
            if (_panes.TryGetValue(wn, out PaneEntry existing)) return existing;

            var ctrl = new IconEditorPaneControl();
            var ctp = this.CustomTaskPanes.Add(ctrl, "Icon Editor Pane", wn);
            ctp.Width = 380;
            ctp.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            ctp.Visible = false;
            ctp.VisibleChanged += (s, args) =>
            {
                EditorPaneVisibilityChanged?.Invoke(this, EventArgs.Empty);
            };

            var entry = new PaneEntry { TaskPane = ctp, Pane = ctrl };
            _panes[wn] = entry;
            return entry;
        }

        private PaneEntry GetOrCreateActivePaneEntry()
        {
            try
            {
                var wn = this.Application.ActiveWindow;
                return EnsurePaneForWindow(wn);
            }
            catch
            {
                return null;
            }
        }

        private void CleanupClosedWindows()
        {
            var alive = new HashSet<PowerPoint.DocumentWindow>();
            foreach (PowerPoint.DocumentWindow w in this.Application.Windows)
            {
                alive.Add(w);
            }
            var dead = _panes.Keys.Where(k => !alive.Contains(k)).ToList();
            foreach (var k in dead)
            {
                try { _panes[k].TaskPane.Dispose(); } catch { }
                _panes.Remove(k);
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new IconBuilderRibbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
