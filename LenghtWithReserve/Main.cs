using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace LenghtWithReserve
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
                if (CreateSharedParameter(application.Application, doc, "Длина с запасом", categorySet, BuiltInParameterGroup.PG_DATA, true))
                {
                    transaction.Commit();
                }
            }
            catch (Exception ex)
            {
                transaction.RollBack();
            }
            transaction.Dispose();
            var selectedObjects = uiDoc.Selection.PickObjects(ObjectType.Element, new PipeFilter(), "Выберите трубы");
            foreach (var obj in selectedObjects)
            {
                using var tr = new Transaction(doc,"Add LenghtWithReserve value");
                tr.Start();
                try
                {
                    var el = (Pipe)doc.GetElement(obj);
                    Parameter lenght = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                    var meterLenght = UnitUtils.ConvertFromInternalUnits(lenght.AsDouble(), UnitTypeId.Meters);
                    Parameter lenghtWithReserve = el.LookupParameter("Длина с запасом");
                    if (lenghtWithReserve == null)
                    {
                        TaskDialog.Show("Info", "Параметр не найден");
                        return Result.Succeeded;
                    }
                    lenghtWithReserve.Set(meterLenght * 1.1);
                    tr.Commit();
                }
                catch (Exception e)
                {
                    tr.RollBack();
                }
            }
            return Result.Succeeded;
        }

        private class PipeFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem)
            {
                return elem is Pipe;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                return true;
            }
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
