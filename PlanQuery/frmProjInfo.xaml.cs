using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace PlanQuery
{
    /// <summary>
    /// Interaction logic for frmProjInfo.xaml
    /// </summary>
    public partial class frmProjInfo : Window
    {
        #region Fields

        private readonly Document _curDoc;

        private static readonly List<string> SpecLevels = new()
        {
            "Standard",
            "Premium",
            "Luxury",
            "Custom"
        };

        private static readonly List<string> ClientNames = new()
        {
            // Add your client names here
            "Client A",
            "Client B",
            "Client C"
        };

        private static readonly List<string> ClientDivisions = new()
        {
            // Add your division names here
            "Division 1",
            "Division 2",
            "Division 3"
        };

        private static readonly List<string> GarageLoadings = new()
        {
            "Front",
            "Side",
            "Rear"
        };

        #endregion

        #region Constructor

        public frmProjInfo(Document curDoc)
        {
            InitializeComponent();
            _curDoc = curDoc;
            PopulateDropdowns();
            LoadExistingValues();
        }

        #endregion

        #region Initialization

        private void PopulateDropdowns()
        {
            cbxSpecLevel.ItemsSource = SpecLevels;
            cbxClientName.ItemsSource = ClientNames;
            cbxClientDivision.ItemsSource = ClientDivisions;
            cbxGarageLoading.ItemsSource = GarageLoadings;
        }

        private void LoadExistingValues()
        {
            ProjectInfo projInfo = _curDoc.ProjectInformation;

            tbxPlanName.Text = Common.Utils.GetParameterValueByName(projInfo, "Project Name") ?? string.Empty;
            tbxClientSubdivision.Text = Common.Utils.GetParameterValueByName(projInfo, "Client Subdivision") ?? string.Empty;

            SetComboValue(cbxSpecLevel, Common.Utils.GetParameterValueByName(projInfo, "Spec Level"));
            SetComboValue(cbxClientName, Common.Utils.GetParameterValueByName(projInfo, "Client Name"));
            SetComboValue(cbxClientDivision, Common.Utils.GetParameterValueByName(projInfo, "Client Division"));
            SetComboValue(cbxGarageLoading, Common.Utils.GetParameterValueByName(projInfo, "Garage Loading"));
        }

        private static void SetComboValue(System.Windows.Controls.ComboBox combo, string value)
        {
            if (string.IsNullOrEmpty(value)) return;

            int index = combo.Items.IndexOf(value);
            if (index >= 0)
                combo.SelectedIndex = index;
            else
                combo.Text = value;
        }

        #endregion

        #region Validation

        private bool ValidateInputs(out string errorMessage)
        {
            errorMessage = string.Empty;
            var missing = new List<string>();

            if (string.IsNullOrWhiteSpace(tbxPlanName.Text)) missing.Add("Plan Name");
            if (string.IsNullOrWhiteSpace(cbxSpecLevel.Text)) missing.Add("Spec Level");
            if (string.IsNullOrWhiteSpace(cbxClientName.Text)) missing.Add("Client Name");
            if (string.IsNullOrWhiteSpace(cbxClientDivision.Text)) missing.Add("Client Division");
            if (string.IsNullOrWhiteSpace(tbxClientSubdivision.Text)) missing.Add("Client Subdivision");
            if (string.IsNullOrWhiteSpace(cbxGarageLoading.Text)) missing.Add("Garage Loading");

            if (missing.Count > 0)
            {
                errorMessage = "The following fields are required:\n\n" +
                               string.Join("\n", missing.Select(f => $"  \u2022 {f}"));
                return false;
            }

            return true;
        }

        #endregion

        #region Write to Revit

        private void WriteValuesToProjectInfo()
        {
            using (Transaction t = new Transaction(_curDoc, "Set Project Information"))
            {
                t.Start();

                ProjectInfo projInfo = _curDoc.ProjectInformation;

                SetParameterValue(projInfo, "Project Name", tbxPlanName.Text.Trim());
                SetParameterValue(projInfo, "Spec Level", cbxSpecLevel.Text.Trim());
                SetParameterValue(projInfo, "Client Name", cbxClientName.Text.Trim());
                SetParameterValue(projInfo, "Client Division", cbxClientDivision.Text.Trim());
                SetParameterValue(projInfo, "Client Subdivision", tbxClientSubdivision.Text.Trim());
                SetParameterValue(projInfo, "Garage Loading", cbxGarageLoading.Text.Trim());

                t.Commit();
            }
        }

        private static void SetParameterValue(ProjectInfo projInfo, string paramName, string value)
        {
            IList<Parameter> paramList = projInfo.GetParameters(paramName);
            if (paramList == null || paramList.Count == 0) return;

            Parameter param = paramList[0];
            if (!param.IsReadOnly)
                param.Set(value);
        }

        #endregion

        #region Event Handlers

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateInputs(out string errorMessage))
            {
                MessageBox.Show(errorMessage, "Missing Information",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                WriteValuesToProjectInfo();
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while saving:\n\n{ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void btnHelp_Click(object sender, RoutedEventArgs e)
        {
            // Add help content here
        }

        #endregion
    }
}
