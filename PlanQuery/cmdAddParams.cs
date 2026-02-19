using PlanQuery.Common;

namespace PlanQuery
{
    [Transaction(TransactionMode.Manual)]
    public class cmdAddParams : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            // set variable for the shared parameter file path and group name
            string sharedParamFile = @"S:\Shared Folders\Lifestyle USA Design\Library 2026\LD_Shared-Parameters_Master.txt";
            string sharedParamGroup = "Project Information";

            // define list of parameters to add
            List<string> listParams = new List<string> { "Spec Level", "Client Name", "Client Division", "Client Subdivision", "Garage Loading" };

            // wrap operation in a try-catch block to handle any errors that may occur
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
                    if (Utils.DoesProjectParamExist(curDoc, curParam))
                        existingParams.Add(curParam);
                    else
                        addedParams.Add(curParam);
                }

                // if all parameters already exist in the project, notify user then launch form
                if (addedParams.Count == 0)
                {
                    Utils.TaskDialogInformation("Plan Query", "Parameters Exist",
                        $"All parameters already exist in the project:\n{string.Join("\n", existingParams)}");
                }
                else
                {
                    // set the shared parameter file and open it
                    uiapp.Application.SharedParametersFilename = sharedParamFile;

                    // set the definition file variable to the currently open shared parameter file
                    DefinitionFile curDefFile = uiapp.Application.OpenSharedParameterFile();

                    // null check the definition file
                    if (curDefFile == null)
                    {
                        Utils.TaskDialogError("Plan Query", "Error",
                            $"Could not open shared parameter file:\n{sharedParamFile}");
                        return Result.Failed;
                    }

                    // set up binding to the Project Information category
                    CategorySet catSet = new CategorySet();
                    Category catProjInfo = curDoc.Settings.Categories.get_Item(BuiltInCategory.OST_ProjectInformation);
                    catSet.Insert(catProjInfo);
                    InstanceBinding instBinding = uiapp.Application.Create.NewInstanceBinding(catSet);

                    // create a transaction to add the parameters to the project
                    using (Transaction t = new Transaction(curDoc, "Add Shared Parameters"))
                    {
                        t.Start();

                        foreach (string curParamName in addedParams)
                        {
                            // get the parameter definition from the shared parameter file
                            Definition curDef = Utils.GetParameterDefinitionFromFile(curDefFile, sharedParamGroup, curParamName);

                            if (curDef == null)
                            {
                                Utils.TaskDialogError("Plan Query", "Error",
                                    $"Could not find definition for parameter '{curParamName}' in shared parameter file:\n{sharedParamFile}");
                                continue;
                            }

                            // bind the parameter to the Project Information category
                            curDoc.ParameterBindings.Insert(curDef, instBinding);
                        }

                        t.Commit();
                    }

                    // build and show the result message
                    string resultMessage = $"Added {addedParams.Count} parameter(s):\n";
                    foreach (string name in addedParams)
                        resultMessage += $"  - {name}\n";

                    if (existingParams.Count > 0)
                    {
                        resultMessage += $"\n{existingParams.Count} parameter(s) already exist in the project and were not added:\n";
                        foreach (string name in existingParams)
                            resultMessage += $"  - {name}\n";
                    }

                    Utils.TaskDialogInformation("Plan Query", "Parameters Added", resultMessage);
                }

                // launch frmProjInfo to collect values — runs whether params were just added or already existed
                frmProjInfo form = new frmProjInfo(curDoc);
                bool? formResult = form.ShowDialog();

                if (formResult != true)
                    return Result.Cancelled;

                // write the collected values to Project Information
                using (Transaction t = new Transaction(curDoc, "Set Project Information"))
                {
                    t.Start();

                    ProjectInfo projInfo = curDoc.ProjectInformation;

                    Utils.SetParameterByName(projInfo, "Project Name", form.PlanName);
                    Utils.SetParameterByName(projInfo, "Spec Level", form.SpecLevel);
                    Utils.SetParameterByName(projInfo, "Client Name", form.ClientName);
                    Utils.SetParameterByName(projInfo, "Client Division", form.ClientDivision);
                    Utils.SetParameterByName(projInfo, "Client Subdivision", form.ClientSubdivision);
                    Utils.SetParameterByName(projInfo, "Garage Loading", form.GarageLoading);

                    t.Commit();
                }

                Utils.TaskDialogInformation("Plan Query", "Success", "Project Information has been updated.");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Utils.TaskDialogError("Plan Query", "Error", $"An error occurred:\n{ex.Message}");
                return Result.Failed;
            }
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnAddParams";
            string buttonTitle = "Add Parameters";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "Add required PlanQuery shared parameters to Project Information, then set their values.");

            return myButtonData.Data;
        }
    }
}