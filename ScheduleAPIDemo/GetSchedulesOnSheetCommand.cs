#region Namespaces
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

#endregion

namespace ScheduleAPIDemo
{
    [Transaction(TransactionMode.Manual)]
    public class GetSchedulesOnSheetCommand : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp =
                commandData.Application;
            UIDocument uidoc =
                uiapp.ActiveUIDocument;
            Document doc =
                uidoc.Document;


            var viewSheet =
                doc.ActiveView as ViewSheet;

            if (viewSheet == null)
            {
                message =
                    "Выберите вид типа Лист";
                return Result.Failed;
            }

            var schedules =                
                viewSheet
                    .GetSchedules()
                    .ToList();

            StringBuilder sb =
                new StringBuilder();
            sb.AppendLine(string
                .Format("{0} " +
                        "спецификаций на листе [{1} - {2}]",
                        schedules.Count,
                        viewSheet.SheetNumber,
                        viewSheet.Name));

            foreach (var viewSchedule in schedules)
            {
                // Do something
                sb.AppendLine(string
                        .Format("\tId: {0} Name: {1}",
                            viewSchedule.Id,
                            viewSchedule.Name));
            }


            TaskDialog.Show("Спецификации на листе",
                sb.ToString());

            return Result.Succeeded;
        }
    }

    public static class ViewSheetExtensions
    {
        public static IEnumerable<ViewSchedule>
            GetSchedules(this ViewSheet viewSheet)
        {
            var doc = viewSheet.Document;

            FilteredElementCollector collector =
                new FilteredElementCollector(doc, viewSheet.Id);

            var scheduleSheetInstances =
                collector
                    .OfClass(typeof(ScheduleSheetInstance))
                    .ToElements()
                    .OfType<ScheduleSheetInstance>();

            foreach (var scheduleSheetInstance in
                scheduleSheetInstances)
            {
                var scheduleId =
                    scheduleSheetInstance
                        .ScheduleId;
                if (scheduleId == ElementId.InvalidElementId)
                    continue;

                var viewSchedule =
                    doc.GetElement(scheduleId)
                    as ViewSchedule;

                if (viewSchedule != null)
                    yield return viewSchedule;
            }
        }
    }
}
