using Autodesk.AutoCAD.Runtime;

namespace XingManager
{
    public class Commands
    {
        [CommandMethod("XINGFORM")]
        public void ShowForm()
        {
            var app = XingManagerApp.Instance;
            if (app == null)
            {
                return;
            }

            app.ShowPalette();
        }

        [CommandMethod("XINGAPPLY")]
        public void ApplyChanges()
        {
            var app = XingManagerApp.Instance;
            if (app == null)
            {
                return;
            }

            var form = app.GetOrCreateForm();
            form?.ApplyToDrawing();
        }

        [CommandMethod("XINGPAGE")]
        public void GeneratePage()
        {
            var app = XingManagerApp.Instance;
            if (app == null)
            {
                return;
            }

            var form = app.GetOrCreateForm();
            form?.GenerateXingPageFromCommand();
        }

        [CommandMethod("XINGLATROW")]
        public void CreateLatLongRow()
        {
            var app = XingManagerApp.Instance;
            if (app == null)
            {
                return;
            }

            var form = app.GetOrCreateForm();
            form?.CreateLatLongRowFromCommand();
        }

        [CommandMethod("XINGREN")]
        public void Renumber()
        {
            var app = XingManagerApp.Instance;
            if (app == null)
            {
                return;
            }

            var form = app.GetOrCreateForm();
            form?.RenumberSequentiallyFromCommand();
        }
    }
}
