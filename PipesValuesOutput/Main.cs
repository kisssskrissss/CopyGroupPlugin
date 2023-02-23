using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PipesValuesOutput
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uIApplication = commandData.Application;
            var uidoc = uIApplication.ActiveUIDocument;
            var doc = uidoc.Document;

            string pipesInfo = string.Empty;

            var pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();

            string selectedFilePath = String.Empty;
            using (SaveFileDialog dialog = new SaveFileDialog()
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "All Files (*.*)|*.*",
                FileName = "Info.xlsx",
                DefaultExt = ".xlsx"
            })
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = dialog.FileName;
                }
            }

            if (!string.IsNullOrEmpty(selectedFilePath))
            {
                using (var fStream = new FileStream(selectedFilePath,FileMode.OpenOrCreate,FileAccess.Write))
                {
                    IWorkbook workbook = new XSSFWorkbook();
                    ISheet sheet = workbook.CreateSheet("Лист 1");

                    int rowIndex = 0;

                    sheet.SetCellValue(0, 0, "Тип трубы");
                    sheet.SetCellValue(0, 1, "Наружний диаметр");
                    sheet.SetCellValue(0, 2, "Внутренний диаметр");
                    sheet.SetCellValue(0, 3, "Длина");

                    rowIndex++;
                    foreach (Pipe pipe in pipes)
                    {
                        var innerDiam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                        var outerDiam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                        var lenght = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                        sheet.SetCellValue(rowIndex, 0, pipe.PipeType.Name);
                        sheet.SetCellValue(rowIndex, 1, outerDiam.AsDouble());
                        sheet.SetCellValue(rowIndex, 2, innerDiam.AsDouble());
                        sheet.SetCellValue(rowIndex, 3, lenght.AsDouble());
                        rowIndex++;
                    }


                    workbook.Write(fStream, false);
                    workbook.Close();
                }
            }

            return Result.Succeeded;
        }
    }
}
