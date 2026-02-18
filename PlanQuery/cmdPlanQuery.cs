using Autodesk.Revit.DB.Architecture;
using PlanQuery.Common;
using Microsoft.Data.SqlClient;
using System.Data.OleDb;

namespace PlanQuery
{
    /// <summary>
    /// Revit command to extract plan data from active model and save to Access database
    /// Requires: clsPlanData.cs
    /// </summary>
    [Transaction(TransactionMode.Manual)]

    public class cmdPlanQuery : IExternalCommand
    {
        private const string dbFilePath = @"S:\Shared Folders\Lifestyle USA Design\HousePlans.accdb";

        //private const string SqlConnStr =
        //    "Server=tcp:placeholder.database.windows.net,1433;Initial Catalog=PlanQuery;User ID=fake;Password=fake;Encrypt=True;TrustServerCertificate=False;Connection Timeout=5;";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Force SqlClient to use managed networking (avoids native SNI issues inside Revit)
            AppContext.SetSwitch("Switch.Microsoft.Data.SqlClient.UseManagedNetworkingOnWindows", true);

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
                UpsertPlan(planData);

                Utils.TaskDialogInformation(
                    "Plan Query",
                    "Success",
                    $"Saved plan '{planData.PlanName}' (Spec: {planData.SpecLevel}, Subdivision: {planData.Subdivision})."
                );


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
            planData.Client = Utils.GetParameterValueByName(curProjInfo, "Client Name");
            planData.Division = Utils.GetParameterValueByName(curProjInfo, "Client Division");
            planData.Subdivision = Utils.GetParameterValueByName(curProjInfo, "Client Subdivision");

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

        /// <summary>
        /// Get overall building dimensions from bounding box
        /// </summary>
        private void GetBuildingDimensions(Document curDoc, out string width, out string depth)
        {
            width = "0'-0\"";
            depth = "0'-0\"";

            // Get all walls to calculate building extents
            FilteredElementCollector collector = new FilteredElementCollector(curDoc)
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
        private int CountStories(Document curDoc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(curDoc)
                .OfClass(typeof(Level));

            // Filter out levels that are not stories (exclude roof, foundation, etc.)
            var levels = collector.Cast<Level>()
                .Where(l => !l.Name.ToLower().Contains("roof") &&
                           !l.Name.ToLower().Contains("foundation") &&
                           !l.Name.ToLower().Contains("base") &&
                           !l.Name.ToLower().Contains("plate"))
                .ToList();

            return levels.Count;
        }

        /// <summary>
        /// Count bedrooms and bathrooms from room elements
        /// </summary>
        private void GetRoomCounts(Document curDoc, out int bedrooms, out decimal bathrooms)
        {
            bedrooms = 0;
            decimal fullBaths = 0;
            decimal halfBaths = 0;

            FilteredElementCollector collector = new FilteredElementCollector(curDoc)
                .OfClass(typeof(SpatialElement))
                .WhereElementIsNotElementType();

            foreach (Room room in collector.OfType<Room>())
            {
                if (room.Area <= 0) continue;

                string roomName = room.Name.ToLower();

                if (roomName.Contains("bedroom") || roomName.Contains("bed"))
                    bedrooms++;

                if (roomName.Contains("bath"))
                {
                    if (roomName.Contains("powder") || roomName.Contains("half"))
                        halfBaths++;
                    else
                        fullBaths++;
                }
            }

            bathrooms = fullBaths + (halfBaths * 0.5m);
        }

        /// <summary>
        /// Count garage bays
        /// </summary>
        private int CountGarageBays(Document curDoc)
        {
            int garageBays = 0;

            FilteredElementCollector collector = new FilteredElementCollector(curDoc)
                .OfClass(typeof(SpatialElement))
                .WhereElementIsNotElementType();

            foreach (Room room in collector.OfType<Room>())
            {
                if (room.Area <= 0) continue;

                string roomName = room.Name.ToLower();
                if (roomName.Contains("garage"))
                {
                    if (roomName.Contains("three"))
                        garageBays += 3;
                    else if (roomName.Contains("two"))
                        garageBays += 2;
                    else if (roomName.Contains("one"))
                        garageBays += 1;
                }
            }

            return garageBays;
        }

        /// <summary>
        /// Get living area from the Floor Areas schedule
        /// </summary>
        private int GetLivingArea(Document curDoc)
        {
            ViewSchedule schedule = Utils.GetFloorAreaSchedule(curDoc);
            if (schedule == null) return 0;

            TableData tableData = schedule.GetTableData();
            TableSectionData bodyData = tableData.GetSectionData(SectionType.Body);

            int rowCount = bodyData.NumberOfRows;
            int areaCol = bodyData.NumberOfColumns - 1;

            for (int row = 0; row < rowCount; row++)
            {
                string cellName = bodyData.GetCellText(row, 0).Trim();

                if (cellName.Equals("Living", StringComparison.OrdinalIgnoreCase))
                {
                    // 1-story: area is on the same row as "Living"
                    string areaText = bodyData.GetCellText(row, areaCol).Trim();
                    if (!string.IsNullOrEmpty(areaText))
                        return ParseAreaValue(areaText);

                    // Multi-story: scan forward for subtotal row (blank name with area value)
                    for (int subRow = row + 1; subRow < rowCount; subRow++)
                    {
                        string subName = bodyData.GetCellText(subRow, 0).Trim();
                        string subArea = bodyData.GetCellText(subRow, areaCol).Trim();

                        // Hit another section header — stop scanning
                        if (!string.IsNullOrEmpty(subName) && !subName.Contains("Floor"))
                            break;

                        // Subtotal row: blank name with an area value
                        if (string.IsNullOrEmpty(subName) && !string.IsNullOrEmpty(subArea))
                            return ParseAreaValue(subArea);
                    }
                }
            }

            return 0;
        }

        /// <summary>
        /// Get total covered area from the Floor Areas schedule
        /// </summary>
        private int GetTotalArea(Document curDoc)
        {
            ViewSchedule schedule = Utils.GetFloorAreaSchedule(curDoc);
            if (schedule == null) return 0;

            TableData tableData = schedule.GetTableData();
            TableSectionData bodyData = tableData.GetSectionData(SectionType.Body);

            int rowCount = bodyData.NumberOfRows;
            int areaCol = bodyData.NumberOfColumns - 1;

            for (int row = 0; row < rowCount; row++)
            {
                string cellName = bodyData.GetCellText(row, 0).Trim();

                if (cellName.Equals("Total Covered", StringComparison.OrdinalIgnoreCase))
                    return ParseAreaValue(bodyData.GetCellText(row, areaCol).Trim());
            }

            return 0;
        }

        /// <summary>
        /// Parse area text like "2206 SF" to integer
        /// </summary>
        private int ParseAreaValue(string areaText)
        {
            string cleaned = areaText.Replace("SF", "").Trim();

            if (int.TryParse(cleaned, out int result))
                return result;

            return 0;
        }

        #endregion

        #region Database Operations

        private bool PlanExistsInDatabase(string planName, string specLevel)
        {
            string query = "SELECT COUNT(*) FROM HousePlans WHERE PlanName = ? AND SpecLevel = ?";

            using (OleDbConnection conn = new OleDbConnection(GetConnectionString()))
            using (OleDbCommand cmd = new OleDbCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("?", planName);
                cmd.Parameters.AddWithValue("?", specLevel);

                conn.Open();
                int count = (int)cmd.ExecuteScalar();
                return count > 0;
            }
        }

        private void UpdatePlanInDatabase(clsPlanData planData)
        {
            string query = @"
        UPDATE HousePlans 
        SET Client = ?, Division = ?, Subdivision = ?,
            OverallWidth = ?, OverallDepth = ?, Stories = ?, 
            Bedrooms = ?, Bathrooms = ?, GarageBays = ?, 
            LivingArea = ?, TotalArea = ?
        WHERE PlanName = ? AND SpecLevel = ?";

            using (OleDbConnection conn = new OleDbConnection(GetConnectionString()))
            using (OleDbCommand cmd = new OleDbCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("?", planData.Client);
                cmd.Parameters.AddWithValue("?", planData.Division);
                cmd.Parameters.AddWithValue("?", planData.Subdivision);
                cmd.Parameters.AddWithValue("?", planData.OverallWidth);
                cmd.Parameters.AddWithValue("?", planData.OverallDepth);
                cmd.Parameters.AddWithValue("?", planData.Stories);
                cmd.Parameters.AddWithValue("?", planData.Bedrooms);
                cmd.Parameters.AddWithValue("?", planData.Bathrooms);
                cmd.Parameters.AddWithValue("?", planData.GarageBays);
                cmd.Parameters.AddWithValue("?", planData.LivingArea);
                cmd.Parameters.AddWithValue("?", planData.TotalArea);
                cmd.Parameters.AddWithValue("?", planData.PlanName);
                cmd.Parameters.AddWithValue("?", planData.SpecLevel);

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        private void InsertPlanIntoDatabase(clsPlanData planData)
        {
            string query = @"
        INSERT INTO HousePlans 
        (PlanName, SpecLevel, Client, Division, Subdivision, OverallWidth, OverallDepth, Stories, Bedrooms, Bathrooms, GarageBays, LivingArea, TotalArea)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

            using (OleDbConnection conn = new OleDbConnection(GetConnectionString()))
            using (OleDbCommand cmd = new OleDbCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("?", planData.PlanName);
                cmd.Parameters.AddWithValue("?", planData.SpecLevel);
                cmd.Parameters.AddWithValue("?", planData.Client);
                cmd.Parameters.AddWithValue("?", planData.Division);
                cmd.Parameters.AddWithValue("?", planData.Subdivision);
                cmd.Parameters.AddWithValue("?", planData.OverallWidth);
                cmd.Parameters.AddWithValue("?", planData.OverallDepth);
                cmd.Parameters.AddWithValue("?", planData.Stories);
                cmd.Parameters.AddWithValue("?", planData.Bedrooms);
                cmd.Parameters.AddWithValue("?", planData.Bathrooms);
                cmd.Parameters.AddWithValue("?", planData.GarageBays);
                cmd.Parameters.AddWithValue("?", planData.LivingArea);
                cmd.Parameters.AddWithValue("?", planData.TotalArea);

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        private string GetConnectionString()
        {
            return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbFilePath};Persist Security Info=False;";
        }

        #endregion

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
