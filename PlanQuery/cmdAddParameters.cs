using PlanQuery.Common;

namespace PlanQuery
{
    [Transaction(TransactionMode.Manual)]
    public class cmdAddParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            // set variable for the shared parameter file path
            string sharedParamFile = @"S:\Shared Folders\Lifestyle USA Design\Library 2026\LD_Shared-Parameters_Master.txt";

            // define list of parameters to add
            List<string> listParams = new List<string> { "Spec Level", "Client Division", "Client Subdivision", "Garage Loading" };


            // add the parameters to the project file
            try
            {
                // check if the shared parameter file exists
                if (!File.Exists(sharedParamFile))
                {
                    Utils.TaskDialogError("Plan Query", "Error",
                        $"Shared parameter file not found:\n{sharedParamFile}");
                    return Result.Failed;
                }

                // define lists for tracking which parameters were added and which already exist in the project
                List<string> addedParams = new List<string>();
                List<string> existingParams = new List<string>();

                // loop through the list of parameters and sort them into the added and existing lists
                foreach (string curParam in listParams)
                {
                    if (Utils.DoesParameterExists(curDoc, curParam))
                    {
                        existingParams.Add(curParam);
                    }
                    else
                    {
                        addedParams.Add(curParam);
                    }
                }








            }
            catch (Exception ex)
            {
                message = $"Error adding parameters: {ex.Message}";
                return Result.Failed;
            }


            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Button 2";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 2");

            return myButtonData.Data;
        }
    }
}
