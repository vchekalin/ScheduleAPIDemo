using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ScheduleAPIDemo
{
    [Transaction(TransactionMode.ReadOnly)]
    public class GetElementsInScheduleCommand : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp =
               commandData.Application;
            UIDocument uidoc =
                uiapp.ActiveUIDocument;
            Document doc =
                uidoc.Document;


            var viewSchedule =
                doc.ActiveView as ViewSchedule;

            if (viewSchedule == null)
            {
                message =
                    "Выберите спецификацию в диспетчере проекта";
                return Result.Failed;
            }

            var elementIds =
                viewSchedule
                    .GetElementIdsInSchedule()
                    .ToList();

            StringBuilder sb =
                new StringBuilder();
            sb.AppendLine(string
                .Format("{0} " +
                        "элементов в спецификации [{1}]",
                        elementIds.Count,
                        viewSchedule.Name));

            foreach (var elementId in elementIds)
            {
                // Если мы выбрали спецификацию типа Расход материала
                // то наравне с элементами спецификации мы получим список всех материалов проекта.
                // Так как нам это не нужно - пропускаем материалы.
                var element =
                    doc.GetElement(elementId);

                if (element is Material)
                    continue;

               
                sb.AppendLine(string
                .Format("\tId: {0} Имя: {1} Тип: {2}",
                    element.Id,
                    element.Name,
                    element.GetType()));
            }


            TaskDialog.Show("Элементы в спецификации",
                sb.ToString());

            return Result.Succeeded;

        }
    }

    public static class ViewScheduleExtensions
    {
        public static IEnumerable<ElementId>
            GetElementIdsInSchedule(this ViewSchedule viewSchedule)
        {
            var doc = viewSchedule.Document;

            FilteredElementCollector collector =
                new FilteredElementCollector(doc, viewSchedule.Id);

            var elementIds =
                collector
                    .WhereElementIsNotElementType()
                    .ToElementIds();

            return elementIds;
        }
    }
}