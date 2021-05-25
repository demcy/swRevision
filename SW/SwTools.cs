using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using SldWorks;

namespace SW
{
    public class SwTools
    {
        private SldWorks.SldWorks _swApp;
        private SldWorks.DrawingDoc _swDrawingDoc;
        private SldWorks.ModelDoc2 _swModel;
        public bool SwConnect()
        {
            try
            {
                _swApp = (SldWorks.SldWorks) Marshal.GetActiveObject("SldWorks.Application");
                Console.WriteLine("SolidWorks is connected");
                return true;
            }
            catch
            {
                Console.WriteLine("SolidWorks not connected");
                return false;
            }
        }
        public void EasyOpen(string name)
        {
            _swDrawingDoc = (SldWorks.DrawingDoc)_swApp.OpenDoc(name, 3);
            _swModel = (SldWorks.ModelDoc2) _swApp.ActiveDoc;
        }
        public void AddRevision(List<string> props)
        {
            var sheet = (SldWorks.Sheet)_swDrawingDoc.GetCurrentSheet();
            var revisionTable = (SldWorks.RevisionTableAnnotation) sheet.RevisionTable;
            revisionTable.AddRevision(props[0]);
            var table = (TableAnnotation)revisionTable;
            if (props[1] != "")
            {
                table.Text[0, 1] = props[1];
            }
            table.Text[0, 2] = props[2];
            int colCount = table.ColumnCount;
            for (int i = 3; i < colCount; i++)
            {
                if (props[i] == "")
                {
                    table.Text[0, i] = table.Text[1, i];
                }
                else
                {
                    table.Text[0, i] = props[i];
                }
            }
            // table.Text[0, 2] = props[2];
            // table.Text[0, 3] = table.Text[1, 3];
            // table.Text[0, 4] = table.Text[1, 4];
            // table.Text[0, 5] = table.Text[1, 5];
            // table.Text[0, 6] = table.Text[1, 6];
        }
        public void SaveToPdf(string n)
        {
            _swDrawingDoc.ForceRebuild();
            var swExportPdfData = (ExportPdfData) _swApp.GetExportFileData(1);
            _swModel.Extension.SaveAs(n, 0, 1, swExportPdfData, 0, 0);
        }
        public void CloseDoc(string name)
        {
            _swModel.Save();
            _swApp.CloseDoc(name);
        }
        public void GetFiles(string[] names, string path, List<string> props)
        {
            if (!SwConnect()) return;
            foreach (var name in names)
            {
                if (CheckExist(name, path, props[0])) continue;
                EasyOpen(name);
                AddRevision(props);
                SaveToPdf(GetName(name, path, props[0]));
                CloseDoc(name);
            }
        }
        public bool CheckExist(string name, string path, string rev)
        {
            string[] p = Directory.GetFiles(path);
            var backIndex = name.LastIndexOf("\\") + 1;
            var pointIndex = name.LastIndexOf(".");
            name = name.Substring(backIndex, pointIndex - backIndex) + "_rev."+rev;
            return p.Any(v => v.Contains(name));
        }
        public string GetName(string name, string path, string rev)
        {
            var backIndex = name.LastIndexOf("\\") + 1;
            var pointIndex = name.LastIndexOf(".");
            name = name.Substring(backIndex, pointIndex - backIndex);
            name = path + name + "_rev."+rev+".pdf";
            return name;
        }
    }
}
