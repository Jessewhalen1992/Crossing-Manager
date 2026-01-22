using System;
using System.Drawing;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using XingManager.Services;

namespace XingManager
{
    public class XingManagerApp : IExtensionApplication
    {
        private static readonly Guid PaletteGuid = new Guid("71E8DF88-8F04-4D7E-AD5F-97F1F4F0F5BB");

        private PaletteSet _palette;
        private XingForm _form;
        private Document _formDocument;
        private readonly TableFactory _tableFactory = new TableFactory();
        private readonly LayoutUtils _layoutUtils = new LayoutUtils();
        private readonly Serde _serde = new Serde();
        private readonly DuplicateResolver _duplicateResolver = new DuplicateResolver();
        private readonly LatLongDuplicateResolver _latLongDuplicateResolver = new LatLongDuplicateResolver();
        private TableSync _tableSync;

        public static XingManagerApp Instance { get; private set; }

        public void Initialize()
        {
            Instance = this;
            _tableSync = new TableSync(_tableFactory);
            Application.DocumentManager.DocumentActivated += OnDocumentActivated;
        }

        public void Terminate()
        {
            Application.DocumentManager.DocumentActivated -= OnDocumentActivated;
            if (_palette != null)
            {
                _palette.Visible = false;
                if (_palette.Count > 0)
                {
                    _palette.Remove(0);
                }

                _form?.Dispose();
                _palette.Dispose();
                _palette = null;
            }

            _form = null;
            _formDocument = null;
            Instance = null;
        }

        private void OnDocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            if (_palette == null || !_palette.Visible)
            {
                return;
            }

            _form = null;
            ShowPalette();
        }

        internal void ShowPalette()
        {
            var form = GetOrCreateForm();
            if (form == null)
            {
                return;
            }

            EnsurePalette();
            ActivatePalette();
        }

        internal XingForm GetOrCreateForm()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return null;
            }

            if (_form == null || _formDocument != doc)
            {
                _form = CreateForm(doc);
                _formDocument = doc;
            }

            return _form;
        }

        private XingForm CreateForm(Document doc)
        {
            var repository = new XingRepository(doc);
            var form = new XingForm(
                doc,
                repository,
                _tableSync,
                _layoutUtils,
                _tableFactory,
                _serde,
                _duplicateResolver,
                _latLongDuplicateResolver);
            form.LoadData();
            AttachForm(form);
            return form;
        }

        private void AttachForm(XingForm form)
        {
            EnsurePalette();
            if (_palette.Count > 0)
            {
                _palette.Remove(0);
                _form?.Dispose();
            }

            _palette.Add("Crossings", form);
            _palette.Visible = true;
            if (_palette.Count > 0)
            {
                _palette.Activate(0);
            }
        }

        private void EnsurePalette()
        {
            if (_palette != null)
            {
                return;
            }

            _palette = new PaletteSet("Crossing Manager", PaletteGuid)
            {
                Style = PaletteSetStyles.ShowAutoHideButton | PaletteSetStyles.ShowCloseButton,
                MinimumSize = new Size(500, 400)
            };
        }

        private void ActivatePalette()
        {
            if (_palette == null)
            {
                return;
            }

            _palette.Visible = true;
            if (_palette.Count > 0)
            {
                _palette.Activate(0);
            }
        }
    }
}

