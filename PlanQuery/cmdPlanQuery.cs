using PlanQuery.Common;
using System.Drawing.Text;
using System.Net.NetworkInformation;
using System.Windows.Controls;

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
                if (!ShowConfirmationDialog(planData))
                {
                    return Result.Cancelled;
                }

                // check if plan already exists in the database before saving
                if (PlanExistsInDatabase(planData.PlanName, planData.SpecLevel))
                {
                    string existsMessage = $"Plan '{planData.PlanName}' with spec '{planData.SpecLevel}' already exists.\n\nDo you want to update it?";

                    // prompt user to confirm if they want to update the existing plan or cancel the operation
                    if (!Utils.TaskDialogAccept("Plan Query", "Plan Exists", existsMessage))
                    {
                        // if user selects no cancel the operation
                        return Result.Cancelled;
                    }

                    // if user selects yes update the existing plan in the database
                    UpdatePlanInDatabase(planData);
                    Utils.TaskDialogInformation("Plan Query", "Success", $"Updated plan '{planData.PlanName}' in database.");
                }
                else
                {
                    // if the plan does not exist in the database insert it as a new record
                    InsertPlanIntoDatabase(planData);
                    Utils.TaskDialogInformation("cmdPlanQuery", "Success", $"Added plan '{planData.PlanName}' to database.");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Utils.TaskDialogError("cmdPlanQuery", "Error", $"An error occurred:\n{ex.Message}");
                return Result.Failed;
            }
        }

        #region Plan Data Extraction

        /// <summary>
        /// Extract plan data from the active Revit model
        /// </summary>
        private clsPlanData ExtractPlanData(Document curDoc)
        {
            var planData = new clsPlanData();

            // create variable for project information
            ProjectInfo curProjInfo = curDoc.ProjectInformation;

            // get plan data from project information and store it
            planData.PlanName = Utils.GetParameterValueByName(curProjInfo, "Project Name");

            planData.SpecLevel = Utils.GetParameterValueByName(curProjInfo, "Spec Level");
            planData.Client = Utils.GetParameterValueByName(curProjInfo, "Client");
            planData.Division = Utils.GetParameterValueByName(curProjInfo, "Division");
            planData.Subdivision = Utils.GetParameterValueByName(curProjInfo, "Subdivision");

            // Get overall building dimensions
            GetBuildingDimensions(curDoc, out string width, out string depth);
            planData.OverallWidth = width;
            planData.OverallDepth = depth;

            // Count stories (levels)
            planData.Stories = CountStories(curDoc);

            // Count bedrooms and bathrooms
            GetRoomCounts(curDoc, out int bedrooms, out decimal bathrooms);
            planData.Bedrooms = bedrooms;
            planData.Bathrooms = bathrooms;

            // Count garage bays
            planData.GarageBays = CountGarageBays(curDoc);

            // Calculate areas
            planData.LivingArea = GetLivingArea(curDoc);
            planData.TotalArea = GetTotalArea(curDoc);

            return planData;
        }

        private void GetBuildingDimensions(Document curDoc, out string width, out string depth)
        {
            throw new NotImplementedException();
        }

        private int CountStories(Document curDoc)
        {
            throw new NotImplementedException();
        }

        private void GetRoomCounts(Document curDoc, out int bedrooms, out decimal bathrooms)
        {
            throw new NotImplementedException();
        }

        private int CountGarageBays(Document curDoc)
        {
            throw new NotImplementedException();
        }       

        private int GetTotalArea(Document curDoc)
        {
            throw new NotImplementedException();
        }

        private int GetLivingArea(Document curDoc)
        {
            throw new NotImplementedException();
        }        

        #endregion

        private void InsertPlanIntoDatabase(clsPlanData planData)
        {
            throw new NotImplementedException();
        }

        private void UpdatePlanInDatabase(clsPlanData planData)
        {
            throw new NotImplementedException();
        }

        private bool PlanExistsInDatabase(string planName, string specLevel)
        {
            throw new NotImplementedException();
        }

        private bool ShowConfirmationDialog(clsPlanData planData)
        {
            string message = $@"Ready to save this plan to the database:

                Plan Name: {planData.PlanName}
                Spec Level: {planData.SpecLevel}
                Client: {planData.Client}
                Division: {planData.Division}
                Subdivision: {planData.Subdivision}

                Dimensions: {planData.OverallWidth} W x {planData.OverallDepth} D
                Total Area: {planData.TotalArea:N0} SF
                Living Area: {planData.LivingArea:N0} SF
                Bedrooms: {planData.Bedrooms}
                Bathrooms: {planData.Bathrooms}
                Stories: {planData.Stories}
                Garage Bays: {planData.GarageBays}

                Do you want to proceed?";

            return Utils.TaskDialogAccept("Plan Query", "Confirm Plan Data", message);
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
