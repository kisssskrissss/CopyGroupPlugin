using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProjectParameter
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication application = commandData.Application;
            UIDocument uiDoc = application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            var categorySet = new CategorySet();
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves));
            using var transaction = new Transaction(doc, "Add param");
            transaction.Start();
            try
            {
                if (CreateSharedParameter(application.Application, doc, "Диаметры трубы", categorySet, BuiltInParameterGroup.PG_TEXT, true))
                {
                    transaction.Commit();
                }
            }
            catch (Exception ex)
            {
                transaction.RollBack();
            }
            transaction.Dispose();

            List<Pipe> pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                .ToList();
            foreach (var pipe in pipes)
            {
                using var tr = new Transaction(doc, "Set param");
                tr.Start();
                try
                {
                    var innerDiam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                    var innerDiamInMm = UnitUtils.ConvertFromInternalUnits(innerDiam.AsDouble(), UnitTypeId.Millimeters);
                    var outerDiam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                    var outerDiamInMm = UnitUtils.ConvertFromInternalUnits(innerDiam.AsDouble(), UnitTypeId.Millimeters);
                    var param = pipe.LookupParameter("Диаметры трубы");
                    param.Set($"Труба {outerDiamInMm}/{innerDiamInMm}");
                    tr.Commit();
                }
                catch
                {
                    tr.RollBack();
                }
            }

            return Result.Succeeded;
        }

        private bool CreateSharedParameter(Application application, Document doc, string parameterName, CategorySet categorySet, BuiltInParameterGroup builtInParameterGroup, bool isInstance)
        {
            DefinitionFile definitionFile = application.OpenSharedParameterFile();
            if (definitionFile == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл общих параметров");
                throw new Exception();
            }

            Definition definition = definitionFile.Groups.SelectMany(group => group.Definitions)
                .FirstOrDefault(def => def.Name == parameterName);
            if (definition == null)
            {
                TaskDialog.Show("Ошибка", $"Не найден параметр с именем {parameterName}");
                throw new Exception();
            }

            Binding binding = null;
            if (isInstance)
            {
                binding = application.Create.NewInstanceBinding(categorySet);
            }
            else
            {
                binding = application.Create.NewTypeBinding(categorySet);
            }

            BindingMap map = doc.ParameterBindings;
            return map.Insert(definition, binding, builtInParameterGroup);
        }
    }
}
