using Autodesk.Revit.DB.Architecture;
using PlanQuery.Common;
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
        private const string DbFilePath = @"S:\Shared Folders\Lifestyle USA Design\HousePlans.accdb";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            try
            {
                // Extract plan data from the active Revit model
                clsPlanData planData = ExtractPlanData(curDoc);

                if (planData == null)
                {
                    Utils.TaskDialogError("Plan Query", "Error", "Unable to extract plan data from the model.");
                    return Result.Failed;
                }

                // Validate required fields before saving
                if (string.IsNullOrWhiteSpace(planData.PlanName) ||
                    string.IsNullOrWhiteSpace(planData.SpecLevel) ||
                    string.IsNullOrWhiteSpace(planData.Client) ||
                    string.IsNullOrWhiteSpace(planData.Division) ||
                    string.IsNullOrWhiteSpace(planData.Subdivision))
                {
                    Utils.TaskDialogWarning("Plan Query", "Missing Information",
                        "All of the following fields are required:\n\n" +
                        "- Plan Name\n- Spec Level\n- Client\n- Division\n- Subdivision\n\n" +
                        "Please set these in Project Information.");
                    return Result.Failed;
                }

                // Show confirmation dialog before saving
                if (!ShowConfirmationDialog(planData))
                    return Result.Cancelled;

                // Check if plan already exists in the database
                if (PlanExistsInDatabase(planData.PlanName, planData.SpecLevel, planData.Subdivision))
                {
                    string existsMessage =
                        $"Plan '{planData.PlanName}' with spec '{planData.SpecLevel}' already exists " +
                        $"in subdivision '{planData.Subdivision}'.\n\nDo you want to update it?";

                    if (!Utils.TaskDialogAccept("Plan Query", "Plan Exists", existsMessage))
                        return Result.Cancelled;

                    UpdatePlanInDatabase(planData);
                    Utils.TaskDialogInformation("Plan Query", "Success",
                        $"Updated plan '{planData.PlanName}' in database.");
                }
                else
                {
                    InsertPlanIntoDatabase(planData);
                    Utils.TaskDialogInformation("Plan Query", "Success",
                        $"Added plan '{planData.PlanName}' to database.");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Utils.TaskDialogError("Plan Query", "Error", $"An error occurred:\n{ex.Message}");
                return Result.Failed;
            }
        }

        #region Plan Data Extraction

        private clsPlanData ExtractPlanData(Document curDoc)
        {
            var planData = new clsPlanData();
            ProjectInfo curProjInfo = curDoc.ProjectInformation;

            planData.PlanName = Utils.GetParameterValueByName(curProjInfo, "Project Name");
            planData.SpecLevel = Utils.GetParameterValueByName(curProjInfo, "Spec Level");
            planData.Client = Utils.GetParameterValueByName(curProjInfo, "Client");
            planData.Division = Utils.GetParameterValueByName(curProjInfo, "Division");
            planData.Subdivision = Utils.GetParameterValueByName(curProjInfo, "Subdivision");

            GetBuildingDimensions(curDoc, out string width, out string depth);
            planData.OverallWidth = width;
            planData.OverallDepth = depth;

            planData.Stories = CountStories(curDoc);

            GetRoomCounts(curDoc, out int bedrooms, out decimal bathrooms);
            planData.Bedrooms = bedrooms;
            planData.Bathrooms = bathrooms;

            planData.GarageBays = CountGarageBays(curDoc);
            planData.LivingArea = GetLivingArea(curDoc);
            planData.TotalArea = GetTotalArea(curDoc);

            return planData;
        }

        private void GetBuildingDimensions(Document curDoc, out string width, out string depth)
        {
            width = "0'-0\"";
            depth = "0'-0\"";

            FilteredElementCollector collector = new FilteredElementCollector(curDoc)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType();

            if (!collector.Any()) return;

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
                    bbox.Min = new XYZ(Math.Min(bbox.Min.X, wallBox.Min.X),
                                       Math.Min(bbox.Min.Y, wallBox.Min.Y),
                                       Math.Min(bbox.Min.Z, wallBox.Min.Z));
                    bbox.Max = new XYZ(Math.Max(bbox.Max.X, wallBox.Max.X),
                                       Math.Max(bbox.Max.Y, wallBox.Max.Y),
                                       Math.Max(bbox.Max.Z, wallBox.Max.Z));
                }
            }

            if (bbox != null)
            {
                width = FormatDimension(bbox.Max.X - bbox.Min.X);
                depth = FormatDimension(bbox.Max.Y - bbox.Min.Y);
            }
        }

        private string FormatDimension(double decimalFeet)
        {
            int feet = (int)Math.Floor(decimalFeet);
            double totalInches = (decimalFeet - feet) * 12.0;
            int inches = (int)Math.Floor(totalInches);
            double remainingInches = totalInches - inches;
            string fraction = "";

            if (remainingInches >= 0.25 && remainingInches < 0.75)
                fraction = " 1/2";
            else if (remainingInches >= 0.75)
                inches++;

            if (inches >= 12) { feet++; inches -= 12; }

            return $"{feet}'-{inches}{fraction}\"";
        }

        private int CountStories(Document curDoc)
        {
            return new FilteredElementCollector(curDoc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .Where(l => !l.Name.ToLower().Contains("roof") &&
                            !l.Name.ToLower().Contains("foundation") &&
                            !l.Name.ToLower().Contains("base") &&
                            !l.Name.ToLower().Contains("plate"))
                .Count();
        }

        private void GetRoomCounts(Document curDoc, out int bedrooms, out decimal bathrooms)
        {
            bedrooms = 0;
            decimal fullBaths = 0;
            decimal halfBaths = 0;

            foreach (Room room in new FilteredElementCollector(curDoc)
                .OfClass(typeof(Room)).Cast<Room>())
            {
                if (room.Area <= 0) continue;

                string name = room.Name.ToLower();

                if (name.Contains("bedroom") || name.Contains("bed"))
                    bedrooms++;

                if (name.Contains("bath"))
                {
                    if (name.Contains("powder") || name.Contains("half"))
                        halfBaths++;
                    else
                        fullBaths++;
                }
            }

            bathrooms = fullBaths + (halfBaths * 0.5m);
        }

        private int CountGarageBays(Document curDoc)
        {
            int bays = 0;

            foreach (Room room in new FilteredElementCollector(curDoc)
                .OfClass(typeof(Room)).Cast<Room>())
            {
                if (room.Area <= 0) continue;

                string name = room.Name.ToLower();
                if (!name.Contains("garage")) continue;

                if (name.Contains("three")) bays += 3;
                else if (name.Contains("two")) bays += 2;
                else if (name.Contains("one")) bays += 1;
            }

            return bays;
        }

        private int GetLivingArea(Document curDoc)
        {
            ViewSchedule schedule = Utils.GetFloorAreaSchedule(curDoc);
            if (schedule == null) return 0;

            TableSectionData body = schedule.GetTableData().GetSectionData(SectionType.Body);
            int rowCount = body.NumberOfRows;
            int areaCol = body.NumberOfColumns - 1;

            for (int row = 0; row < rowCount; row++)
            {
                if (!body.GetCellText(row, 0).Trim().Equals("Living", StringComparison.OrdinalIgnoreCase))
                    continue;

                // 1-story: area is on the same row as "Living"
                string areaText = body.GetCellText(row, areaCol).Trim();
                if (!string.IsNullOrEmpty(areaText)) return ParseAreaValue(areaText);

                // Multi-story: scan forward for subtotal row
                for (int sub = row + 1; sub < rowCount; sub++)
                {
                    string subName = body.GetCellText(sub, 0).Trim();
                    string subArea = body.GetCellText(sub, areaCol).Trim();

                    if (!string.IsNullOrEmpty(subName) && !subName.Contains("Floor")) break;
                    if (string.IsNullOrEmpty(subName) && !string.IsNullOrEmpty(subArea))
                        return ParseAreaValue(subArea);
                }
            }

            return 0;
        }

        private int GetTotalArea(Document curDoc)
        {
            ViewSchedule schedule = Utils.GetFloorAreaSchedule(curDoc);
            if (schedule == null) return 0;

            TableSectionData body = schedule.GetTableData().GetSectionData(SectionType.Body);
            int rowCount = body.NumberOfRows;
            int areaCol = body.NumberOfColumns - 1;

            for (int row = 0; row < rowCount; row++)
            {
                if (body.GetCellText(row, 0).Trim().Equals("Total Covered", StringComparison.OrdinalIgnoreCase))
                    return ParseAreaValue(body.GetCellText(row, areaCol).Trim());
            }

            return 0;
        }

        private int ParseAreaValue(string areaText)
        {
            string cleaned = areaText.Replace("SF", "").Trim();
            return int.TryParse(cleaned, out int result) ? result : 0;
        }

        #endregion

        #region Database Operations

        private bool PlanExistsInDatabase(string planName, string specLevel, string subdivision)
        {
            string query = "SELECT COUNT(*) FROM HousePlans " +
                           "WHERE PlanName = ? AND SpecLevel = ? AND Subdivision = ?";

            using (OleDbConnection conn = new OleDbConnection(GetConnectionString()))
            using (OleDbCommand cmd = new OleDbCommand(query, conn))
            {
                // Positional order must match the ? placeholders in the query above
                cmd.Parameters.AddWithValue("@PlanName", planName);
                cmd.Parameters.AddWithValue("@SpecLevel", specLevel);
                cmd.Parameters.AddWithValue("@Subdivision", subdivision);

                conn.Open();
                int count = (int)cmd.ExecuteScalar();
                return count > 0;
            }
        }

        private void InsertPlanIntoDatabase(clsPlanData plan)
        {
            string query = @"
                INSERT INTO HousePlans 
                    (PlanName, SpecLevel, Client, Division, Subdivision,
                     OverallWidth, OverallDepth, Stories, Bedrooms, Bathrooms,
                     GarageBays, LivingArea, TotalArea)
                VALUES 
                    (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

            using (OleDbConnection conn = new OleDbConnection(GetConnectionString()))
            using (OleDbCommand cmd = new OleDbCommand(query, conn))
            {
                // Positional order must match the column list above
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

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        private void UpdatePlanInDatabase(clsPlanData plan)
        {
            string query = @"
                UPDATE HousePlans 
                SET Client        = ?,
                    Division      = ?,
                    OverallWidth  = ?,
                    OverallDepth  = ?,
                    Stories       = ?,
                    Bedrooms      = ?,
                    Bathrooms     = ?,
                    GarageBays    = ?,
                    LivingArea    = ?,
                    TotalArea     = ?
                WHERE PlanName    = ? 
                  AND SpecLevel   = ? 
                  AND Subdivision = ?";

            using (OleDbConnection conn = new OleDbConnection(GetConnectionString()))
            using (OleDbCommand cmd = new OleDbCommand(query, conn))
            {
                // SET fields first, then WHERE fields - must match the order above
                cmd.Parameters.AddWithValue("@Client", plan.Client);
                cmd.Parameters.AddWithValue("@Division", plan.Division);
                cmd.Parameters.AddWithValue("@OverallWidth", plan.OverallWidth);
                cmd.Parameters.AddWithValue("@OverallDepth", plan.OverallDepth);
                cmd.Parameters.AddWithValue("@Stories", plan.Stories);
                cmd.Parameters.AddWithValue("@Bedrooms", plan.Bedrooms);
                cmd.Parameters.AddWithValue("@Bathrooms", plan.Bathrooms);
                cmd.Parameters.AddWithValue("@GarageBays", plan.GarageBays);
                cmd.Parameters.AddWithValue("@LivingArea", plan.LivingArea);
                cmd.Parameters.AddWithValue("@TotalArea", plan.TotalArea);
                // WHERE fields last
                cmd.Parameters.AddWithValue("@PlanName", plan.PlanName);
                cmd.Parameters.AddWithValue("@SpecLevel", plan.SpecLevel);
                cmd.Parameters.AddWithValue("@Subdivision", plan.Subdivision);

                conn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        private string GetConnectionString()
        {
            return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={DbFilePath};Persist Security Info=False;";
        }

        #endregion

        #region UI

        private bool ShowConfirmationDialog(clsPlanData planData)
        {
            string message = $@"Ready to save this plan to the database:

Plan Name:   {planData.PlanName}
Spec Level:  {planData.SpecLevel}
Client:      {planData.Client}
Division:    {planData.Division}
Subdivision: {planData.Subdivision}

Dimensions:  {planData.OverallWidth} W x {planData.OverallDepth} D
Total Area:  {planData.TotalArea:N0} SF
Living Area: {planData.LivingArea:N0} SF
Bedrooms:    {planData.Bedrooms}
Bathrooms:   {planData.Bathrooms}
Stories:     {planData.Stories}
Garage Bays: {planData.GarageBays}

Do you want to proceed?";

            return Utils.TaskDialogAccept("Plan Query", "Confirm Plan Data", message);
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnPlanQuery";
            string buttonTitle = "Plan Query";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "Extract plan data from the active model and save to the Access database");

            return myButtonData.Data;
        }

        #endregion
    }
}