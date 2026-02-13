using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using PlanQuery.Common;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace HousePlansDatabase
{
    /// <summary>
    /// Revit command to extract plan data from active model and save to Access database
    /// Requires: clsPlanData.cs
    /// </summary>
    [Transaction(TransactionMode.ReadOnly)]
    public class cmdPlanQuery : IExternalCommand
    {
        private const string DATABASE_PATH = @"S:\Shared Folders\Lifestyle USA Design\HousePlans.accdb";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // Extract plan data from Revit model (including Project Information)
                clsPlanData planData = ExtractPlanDataFromModel(doc);

                if (planData == null)
                {
                    Utils.TaskDialogError("cmdPlanQuery", "Error", "Unable to extract plan data from the model.");
                    return Result.Failed;
                }

                // Validate required fields
                if (string.IsNullOrWhiteSpace(planData.PlanName) ||
                    string.IsNullOrWhiteSpace(planData.SpecLevel) ||
                    string.IsNullOrWhiteSpace(planData.Client) ||
                    string.IsNullOrWhiteSpace(planData.Division) ||
                    string.IsNullOrWhiteSpace(planData.Subdivision))
                {
                    Utils.TaskDialogWarning("cmdPlanQuery", "Missing Information",
                        "All of the following fields are required:\n\n" +
                        "- Plan Name (Project Name or Building Name)\n" +
                        "- Spec Level\n" +
                        "- Client\n" +
                        "- Division\n" +
                        "- Subdivision\n\n" +
                        "Please set these in Project Information.");
                    return Result.Failed;
                }

                // Show confirmation dialog with extracted data
                if (!ShowConfirmationDialog(planData))
                {
                    return Result.Cancelled;
                }

                // Check if plan already exists
                if (PlanExistsInDatabase(planData.PlanName, planData.SpecLevel))
                {
                    string existsMessage = $"Plan '{planData.PlanName}' with spec '{planData.SpecLevel}' already exists.\n\nDo you want to update it?";

                    if (!Utils.TaskDialogAccept("cmdPlanQuery_Confirm", "Plan Exists", existsMessage))
                    {
                        return Result.Cancelled;
                    }

                    UpdatePlanInDatabase(planData);
                    Utils.TaskDialogInformation("cmdPlanQuery", "Success", $"Updated plan '{planData.PlanName}' in database.");
                }
                else
                {
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

        #region Revit Data Extraction

        /// <summary>
        /// Extract plan data from the active Revit model
        /// </summary>
        private clsPlanData ExtractPlanDataFromModel(Document doc)
        {
            var planData = new PlanData();

            // Get all data from Project Information
            ProjectInfo projInfo = doc.ProjectInformation;
            if (projInfo != null)
            {
                // Get plan name
                planData.PlanName = GetProjectName(doc);

                // Get custom parameters from Project Information
                planData.SpecLevel = GetParameterValue(projInfo, "Spec Level");
                planData.Client = GetParameterValue(projInfo, "Client");
                planData.Division = GetParameterValue(projInfo, "Division");
                planData.Subdivision = GetParameterValue(projInfo, "Subdivision");
            }

            // Get overall building dimensions
            GetBuildingDimensions(doc, out string width, out string depth);
            planData.OverallWidth = width;
            planData.OverallDepth = depth;

            // Count stories (levels)
            planData.Stories = CountStories(doc);

            // Count bedrooms and bathrooms
            GetRoomCounts(doc, out int bedrooms, out decimal bathrooms);
            planData.Bedrooms = bedrooms;
            planData.Bathrooms = bathrooms;

            // Count garage bays
            planData.GarageBays = CountGarageBays(doc);

            // Calculate areas
            planData.LivingArea = CalculateLivingArea(doc);
            planData.TotalArea = CalculateTotalArea(doc);

            return planData;
        }

        /// <summary>
        /// Get a parameter value from Project Information (by name)
        /// </summary>
        private string GetParameterValue(ProjectInfo projInfo, string parameterName)
        {
            // Try to get parameter by name
            Parameter param = projInfo.LookupParameter(parameterName);
            if (param != null && param.HasValue)
            {
                if (param.StorageType == StorageType.String)
                {
                    return param.AsString();
                }
                else if (param.StorageType == StorageType.Integer)
                {
                    return param.AsInteger().ToString();
                }
                else if (param.StorageType == StorageType.Double)
                {
                    return param.AsDouble().ToString();
                }
            }

            return null;
        }

        /// <summary>
        /// Get project name from project information
        /// </summary>
        private string GetProjectName(Document doc)
        {
            // Try to get from Project Information
            ProjectInfo projInfo = doc.ProjectInformation;
            if (projInfo != null)
            {
                Parameter nameParam = projInfo.get_Parameter(BuiltInParameter.PROJECT_NAME);
                if (nameParam != null && !string.IsNullOrEmpty(nameParam.AsString()))
                {
                    return nameParam.AsString();
                }

                // Try building name
                Parameter buildingParam = projInfo.get_Parameter(BuiltInParameter.PROJECT_BUILDING_NAME);
                if (buildingParam != null && !string.IsNullOrEmpty(buildingParam.AsString()))
                {
                    return buildingParam.AsString();
                }
            }

            // Fallback to document title
            return doc.Title.Replace(".rvt", "").Replace(".RVT", "");
        }

        /// <summary>
        /// Get overall building dimensions from bounding box
        /// </summary>
        private void GetBuildingDimensions(Document doc, out string width, out string depth)
        {
            width = "0'-0\"";
            depth = "0'-0\"";

            // Get all walls to calculate building extents
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType();

            if (collector.Count() == 0)
                return;

            BoundingBoxXYZ bbox = null;

            foreach (Wall wall in collector.Cast<Wall>())
            {
                BoundingBoxXYZ wallBox = wall.get_BoundingBox(null);
                if (wallBox == null) continue;

                if (bbox == null)
                {
                    bbox = wallBox;
                }
                else
                {
                    bbox.Min = new XYZ(
                        Math.Min(bbox.Min.X, wallBox.Min.X),
                        Math.Min(bbox.Min.Y, wallBox.Min.Y),
                        Math.Min(bbox.Min.Z, wallBox.Min.Z));
                    bbox.Max = new XYZ(
                        Math.Max(bbox.Max.X, wallBox.Max.X),
                        Math.Max(bbox.Max.Y, wallBox.Max.Y),
                        Math.Max(bbox.Max.Z, wallBox.Max.Z));
                }
            }

            if (bbox != null)
            {
                // Convert from Revit internal units (decimal feet) to formatted dimension
                double widthFeet = bbox.Max.X - bbox.Min.X;
                double depthFeet = bbox.Max.Y - bbox.Min.Y;

                width = FormatDimension(widthFeet);
                depth = FormatDimension(depthFeet);
            }
        }

        /// <summary>
        /// Convert decimal feet to formatted dimension string (e.g., "65'-8 1/2\"")
        /// </summary>
        private string FormatDimension(double decimalFeet)
        {
            // Get whole feet
            int feet = (int)Math.Floor(decimalFeet);

            // Get remaining inches
            double remainingFeet = decimalFeet - feet;
            double totalInches = remainingFeet * 12.0;

            // Get whole inches
            int inches = (int)Math.Floor(totalInches);

            // Get fraction of inch (round to nearest 1/2")
            double remainingInches = totalInches - inches;
            string fraction = "";

            if (remainingInches >= 0.25 && remainingInches < 0.75)
            {
                fraction = " 1/2";
            }
            else if (remainingInches >= 0.75)
            {
                inches++;
            }

            // Handle inch overflow
            if (inches >= 12)
            {
                feet++;
                inches -= 12;
            }

            return $"{feet}'-{inches}{fraction}\"";
        }

        /// <summary>
        /// Count number of stories in the model
        /// </summary>
        private int CountStories(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(Level));

            // Filter out levels that are not stories (exclude roof, foundation, etc.)
            var levels = collector.Cast<Level>()
                .Where(l => !l.Name.ToLower().Contains("roof") &&
                           !l.Name.ToLower().Contains("foundation") &&
                           !l.Name.ToLower().Contains("base"))
                .ToList();

            return levels.Count;
        }

        /// <summary>
        /// Count bedrooms and bathrooms from room elements
        /// </summary>
        private void GetRoomCounts(Document doc, out int bedrooms, out decimal bathrooms)
        {
            bedrooms = 0;
            decimal fullBaths = 0;
            decimal halfBaths = 0;

            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(Room));

            foreach (Room room in collector.Cast<Room>())
            {
                if (room.Area <= 0) continue; // Skip unplaced rooms

                string roomName = room.Name.ToLower();

                // Count bedrooms
                if (roomName.Contains("bedroom") || roomName.Contains("bed"))
                {
                    bedrooms++;
                }

                // Count bathrooms
                if (roomName.Contains("bath"))
                {
                    if (roomName.Contains("powder") || roomName.Contains("half"))
                    {
                        halfBaths++;
                    }
                    else
                    {
                        fullBaths++;
                    }
                }
            }

            bathrooms = fullBaths + (halfBaths * 0.5m);
        }

        /// <summary>
        /// Count garage bays
        /// </summary>
        private int CountGarageBays(Document doc)
        {
            int garageBays = 0;

            // Method 1: Look for rooms named "Garage"
            FilteredElementCollector roomCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(Room));

            foreach (Room room in roomCollector.Cast<Room>())
            {
                if (room.Area <= 0) continue;

                string roomName = room.Name.ToLower();
                if (roomName.Contains("garage"))
                {
                    // Try to extract number from room name like "2 Car Garage"
                    if (roomName.Contains("2") || roomName.Contains("two"))
                        garageBays = Math.Max(garageBays, 2);
                    else if (roomName.Contains("3") || roomName.Contains("three"))
                        garageBays = Math.Max(garageBays, 3);
                    else if (roomName.Contains("1") || roomName.Contains("one") || garageBays == 0)
                        garageBays = Math.Max(garageBays, 1);
                }
            }

            // Method 2: Count garage doors if no room found
            if (garageBays == 0)
            {
                FilteredElementCollector doorCollector = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Doors)
                    .WhereElementIsNotElementType();

                foreach (FamilyInstance door in doorCollector.Cast<FamilyInstance>())
                {
                    string doorType = door.Name.ToLower();
                    if (doorType.Contains("garage"))
                    {
                        garageBays++;
                    }
                }
            }

            return garageBays;
        }

        /// <summary>
        /// Calculate total living area (excluding garage, porches, etc.)
        /// </summary>
        private int CalculateLivingArea(Document doc)
        {
            double totalArea = 0;

            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(Room));

            foreach (Room room in collector.Cast<Room>())
            {
                if (room.Area <= 0) continue; // Skip unplaced rooms

                string roomName = room.Name.ToLower();

                // Exclude non-living spaces
                if (roomName.Contains("garage") ||
                    roomName.Contains("porch") ||
                    roomName.Contains("patio") ||
                    roomName.Contains("deck") ||
                    roomName.Contains("attic") ||
                    roomName.Contains("crawl"))
                {
                    continue;
                }

                totalArea += room.Area;
            }

            // Convert from square feet to integer
            return (int)Math.Round(totalArea);
        }

        /// <summary>
        /// Calculate total area (all rooms including garage, porches, etc.)
        /// </summary>
        private int CalculateTotalArea(Document doc)
        {
            double totalArea = 0;

            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(Room));

            foreach (Room room in collector.Cast<Room>())
            {
                if (room.Area <= 0) continue; // Skip unplaced rooms

                totalArea += room.Area;
            }

            // Convert from square feet to integer
            return (int)Math.Round(totalArea);
        }

        #endregion

        #region Confirmation Dialog

        /// <summary>
        /// Show confirmation dialog with extracted data
        /// </summary>
        private bool ShowConfirmationDialog(PlanData planData)
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

            return Utils.TaskDialogAccept("cmdPlanQuery_ConfirmData", "Confirm Plan Data", message);
        }

        #endregion

        #region Database Operations

        private string GetConnectionString()
        {
            return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={DATABASE_PATH};Persist Security Info=False;";
        }

        /// <summary>
        /// Check if plan already exists in database
        /// </summary>
        private bool PlanExistsInDatabase(string planName, string specLevel)
        {
            string query = "SELECT COUNT(*) FROM HousePlans WHERE PlanName = @PlanName AND SpecLevel = @SpecLevel";

            using (OleDbConnection conn = new OleDbConnection(GetConnectionString()))
            using (OleDbCommand cmd = new OleDbCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@PlanName", planName);
                cmd.Parameters.AddWithValue("@SpecLevel", specLevel);

                conn.Open();
                int count = (int)cmd.ExecuteScalar();
                return count > 0;
            }
        }

        /// <summary>
        /// Insert new plan into database
        /// </summary>
        private void InsertPlanIntoDatabase(PlanData plan)
        {
            string query = @"
                INSERT INTO HousePlans 
                (PlanName, SpecLevel, Client, Division, Subdivision, OverallWidth, OverallDepth, Stories, Bedrooms, Bathrooms, GarageBays, LivingArea, TotalArea)
                VALUES 
                (@PlanName, @SpecLevel, @Client, @Division, @Subdivision, @OverallWidth, @OverallDepth, @Stories, @Bedrooms, @Bathrooms, @GarageBays, @LivingArea, @TotalArea)";

            using (OleDbConnection conn = new OleDbConnection(GetConnectionString()))
            using (OleDbCommand cmd = new OleDbCommand(query, conn))
            {
                AddParametersToCommand(cmd, plan);
                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Update existing plan in database
        /// </summary>
        private void UpdatePlanInDatabase(PlanData plan)
        {
            string query = @"
                UPDATE HousePlans 
                SET Client = @Client, 
                    Division = @Division, 
                    Subdivision = @Subdivision,
                    OverallWidth = @OverallWidth, 
                    OverallDepth = @OverallDepth, 
                    Stories = @Stories, 
                    Bedrooms = @Bedrooms, 
                    Bathrooms = @Bathrooms, 
                    GarageBays = @GarageBays, 
                    LivingArea = @LivingArea,
                    TotalArea = @TotalArea
                WHERE PlanName = @PlanName AND SpecLevel = @SpecLevel";

            using (OleDbConnection conn = new OleDbConnection(GetConnectionString()))
            using (OleDbCommand cmd = new OleDbCommand(query, conn))
            {
                AddParametersToCommand(cmd, plan);
                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Add parameters to SQL command
        /// </summary>
        private void AddParametersToCommand(OleDbCommand cmd, PlanData plan)
        {
            cmd.Parameters.AddWithValue("@PlanName", plan.PlanName);
            cmd.Parameters.AddWithValue("@SpecLevel", plan.SpecLevel);
            cmd.Parameters.AddWithValue("@Client", plan.Client);
            cmd.Parameters.AddWithValue("@Division", plan.Division);
            cmd.Parameters.AddWithValue("@Subdivision", plan.Subdivision);
            cmd.Parameters.AddWithValue("@OverallWidth", plan.OverallWidth);
            cmd.Parameters.AddWithValue("@OverallDepth", plan.OverallDepth);
            cmd.Parameters.AddWithValue("@Stories", plan.Stories);
            cmd.Parameters.AddWithValue("@Bedrooms", plan.Bedrooms);
            cmd.Parameters.AddWithValue("@Bathrooms", plan.Bathrooms);
            cmd.Parameters.AddWithValue("@GarageBays", plan.GarageBays);
            cmd.Parameters.AddWithValue("@LivingArea", plan.LivingArea);
            cmd.Parameters.AddWithValue("@TotalArea", plan.TotalArea);
        }

        #endregion
    }
}