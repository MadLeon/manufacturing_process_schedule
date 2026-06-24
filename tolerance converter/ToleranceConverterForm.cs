using System;
using System.Globalization;
using System.Windows.Forms;

namespace ToleranceConverter
{
    public partial class ToleranceConverterForm : Form
    {
        private ToleranceDataService? _dataService;
        private bool _isUpdatingUnits = false;

        public ToleranceConverterForm()
        {
            InitializeComponent();
            InitializeDataService();
        }

        private void InitializeDataService()
        {
            try
            {
                _dataService = new ToleranceDataService();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to initialize: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Shared KeyPress handler for both dimension inputs.
        /// Filters non-numeric characters and triggers Convert on Enter.
        /// </summary>
        private void TxtDimension_KeyPress(object sender, KeyPressEventArgs e)
        {
            var tb = (TextBox)sender;

            if (e.KeyChar == (char)Keys.Enter)
            {
                e.Handled = true;
                BtnConvert_Click(sender, e);
                return;
            }

            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
                return;
            }

            if (e.KeyChar == '.' && tb.Text.Contains("."))
                e.Handled = true;
        }

        private void TxtDimension_TextChanged(object sender, EventArgs e)
        {
            if (_isUpdatingUnits) return;
            _isUpdatingUnits = true;
            if (double.TryParse(txtDimension.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double mm) && mm > 0)
                txtDimensionInch.Text = (mm / 25.4).ToString("F5");
            else if (string.IsNullOrWhiteSpace(txtDimension.Text))
                txtDimensionInch.Clear();
            _isUpdatingUnits = false;
        }

        private void TxtDimensionInch_TextChanged(object sender, EventArgs e)
        {
            if (_isUpdatingUnits) return;
            _isUpdatingUnits = true;
            if (double.TryParse(txtDimensionInch.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double inch) && inch > 0)
                txtDimension.Text = (inch * 25.4).ToString("F3");
            else if (string.IsNullOrWhiteSpace(txtDimensionInch.Text))
                txtDimension.Clear();
            _isUpdatingUnits = false;
        }

        private void RbType_CheckedChanged(object sender, EventArgs e)
        {
            var rb = (RadioButton)sender;
            if (!rb.Checked) return;
            lblTableName.Text = rbIT12Half.Checked ? "IT12/2" : "H12";
        }

        private ToleranceType GetSelectedType() =>
            rbIT12Half.Checked ? ToleranceType.IT12Half :
            rbInternal.Checked ? ToleranceType.Internal : ToleranceType.External;

        private void BtnConvert_Click(object? sender, EventArgs e)
        {
            lblError.Text = "";
            lblUpperValue.Text = "-";
            lblLowerValue.Text = "-";

            if (string.IsNullOrWhiteSpace(txtDimension.Text))
            {
                lblError.Text = "Please enter a dimension value";
                return;
            }

            if (!double.TryParse(txtDimension.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double dimension))
            {
                lblError.Text = "Invalid dimension value";
                return;
            }

            if (dimension <= 0 || dimension > 500)
            {
                lblError.Text = "Dimension must be between 0 and 500 mm";
                return;
            }

            if (_dataService == null)
            {
                lblError.Text = "Data service not initialized";
                return;
            }

            var type = GetSelectedType();
            var result = _dataService.GetTolerance(dimension, type);

            lblTableName.Text = type == ToleranceType.IT12Half ? "IT12/2" : "H12";

            if (result.HasValue)
            {
                lblUpperValue.Text = FormatTolerance(result.Value.upper);
                lblLowerValue.Text = FormatTolerance(result.Value.lower);
            }
            else
            {
                lblError.Text = "No tolerance found for this dimension";
            }
        }

        private static string FormatTolerance(double valueMm)
        {
            double valueIn = valueMm / 25.4;
            return $"{valueMm.ToString("+0.000;-0.000;0.000")} mm   /   {valueIn.ToString("+0.00000;-0.00000;0.00000")} in";
        }
    }
}
