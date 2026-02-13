using PlanQuery.Common;
using System.Drawing.Text;

namespace PlanQuery
{
    /// <summary>
    /// Revit command to extract plan data from active model and save to Access database
    /// Requires: clsPlanData.cs
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class cmdPlanQuery : IExternalCommand
    {
        //set variable for the file path and name of the Access database
        private const string dbFilePath = @"S:\Shared Folders\Lifestyle USA Design\HousePlans.accdb";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            try
            {
                // extract plan data from the active Revit model and store it
                clsPlanData planData = ExtractPlanData(curDoc);

                // null check for planData before attempting to save to database
                if(planData == null)
                {
                    Utils.TaskDialogError("Plan Query", "Error", "Unable to extract plan data from the model.");
                    return Result.Failed;
                }

                // validate required fields in planData before saving to database
                if (string.IsNullOrWhiteSpace(planData.PlanName) ||
                    string.IsNullOrWhiteSpace(planData.SpecLevel) ||
                    string.IsNullOrWhiteSpace(planData.Client) ||
                    string.IsNullOrWhiteSpace(planData.Division) ||
                    string.IsNullOrWhiteSpace(planData.Subdivision))
                {
                    Utils.TaskDialogWarning("Plan Query", "Missing Information",
                        "All of the following fields are required:\n\n" +
                        "- Plan Name\n" +
                        "- Spec Level\n" +
                        "- Client\n" +
                        "- Division\n" +
                        "- Subdivision\n\n" +
                        "Please set these in Project Information.");
                    return Result.Failed;
                }

                // show confirmation dialog with extracted before saving to database

            }
            catch (Exception)
            {

                throw;
            }

            
            // Your code goes here

            return Result.Succeeded;
        }

        private clsPlanData ExtractPlanData(Document curDoc)
        {
            throw new NotImplementedException();
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData.Data;
        }
    }

}
