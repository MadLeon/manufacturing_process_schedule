using System;
using System.Windows.Forms;

namespace ToleranceConverter
{
    public partial class ToleranceConverterForm : Form
    {
        private ToleranceDataService? _dataService;

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

        private void TxtDimension_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                e.Handled = true;
                BtnConvert_Click(sender, e);
                return;
            }

            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }

            if (e.KeyChar == '.' && txtDimension.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

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

            if (!double.TryParse(txtDimension.Text, out double dimension))
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

            var result = _dataService.GetTolerance(dimension, rbInternal.Checked);

            if (result.HasValue)
            {
                lblUpperValue.Text = result.Value.upper.ToString("+0.000;-0.000;0.000");
                lblLowerValue.Text = result.Value.lower.ToString("+0.000;-0.000;0.000");
            }
            else
            {
                lblError.Text = "No tolerance found for this dimension";
            }
        }
    }
}
