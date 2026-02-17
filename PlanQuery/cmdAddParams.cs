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
            List<string> listParams = new List<string> { "Spec Level", "Client Division", "Client Subdivision", "Garage Loading" };

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
                    {
                        existingParams.Add(curParam);
                    }
                    else
                    {
                        addedParams.Add(curParam);
                    }
                }

                // if all parameters already exist in the project, notify user and exit
                if (addedParams.Count == 0)
                {
                   Utils.TaskDialogInformation("Plan Query", "Parameters Exist",
                        $"All parameters already exist in the project:\n{string.Join("\n", existingParams)}");
                    return Result.Succeeded;
                }

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
                    // start the transaction
                    t.Start();

                    // loop through the list of parameters to add
                    foreach (string curParamName in addedParams)
                    {
                        // get the parameter definition from the shared parameter file
                        Definition curDef = Utils.GetParameterDefinitionFromFile(curDefFile, sharedParamGroup, curParamName);

                        // null check the parameter definition
                        if (curDef == null)
                        {
                            // notify the user if the parameter definition could not be found
                            Utils.TaskDialogError("Plan Query", "Error",
                                $"Could not find definition for parameter '{curParamName}' in shared parameter file:\n{sharedParamFile}");

                            // skip to the next parameter in the list
                            continue;
                        }

                        // bind the parameter under "Other"
                        curDoc.ParameterBindings.Insert(curDef, instBinding);
                    }

                    // commit the transaction
                    t.Commit();
                }

                // build the result message
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
