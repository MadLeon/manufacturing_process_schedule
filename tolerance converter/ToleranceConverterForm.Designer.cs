namespace ToleranceConverter
{
    partial class ToleranceConverterForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TextBox txtDimension;
        private System.Windows.Forms.TextBox txtDimensionInch;
        private System.Windows.Forms.RadioButton rbInternal;
        private System.Windows.Forms.RadioButton rbExternal;
        private System.Windows.Forms.RadioButton rbIT12Half;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Label lblUpperValue;
        private System.Windows.Forms.Label lblLowerValue;
        private System.Windows.Forms.Label lblError;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblDimension;
        private System.Windows.Forms.Label lblDimensionInch;
        private System.Windows.Forms.GroupBox gbType;
        private System.Windows.Forms.Label lblUpper;
        private System.Windows.Forms.Label lblLower;
        private System.Windows.Forms.Label lblDescription;
        private System.Windows.Forms.Label lblTableName;
        private System.Windows.Forms.Label author;
        private System.Windows.Forms.Label label1;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            txtDimension        = new System.Windows.Forms.TextBox();
            txtDimensionInch    = new System.Windows.Forms.TextBox();
            rbInternal          = new System.Windows.Forms.RadioButton();
            rbExternal          = new System.Windows.Forms.RadioButton();
            rbIT12Half          = new System.Windows.Forms.RadioButton();
            btnConvert          = new System.Windows.Forms.Button();
            lblUpperValue       = new System.Windows.Forms.Label();
            lblLowerValue       = new System.Windows.Forms.Label();
            lblError            = new System.Windows.Forms.Label();
            lblTitle            = new System.Windows.Forms.Label();
            lblDimension        = new System.Windows.Forms.Label();
            lblDimensionInch    = new System.Windows.Forms.Label();
            gbType              = new System.Windows.Forms.GroupBox();
            lblUpper            = new System.Windows.Forms.Label();
            lblLower            = new System.Windows.Forms.Label();
            lblDescription      = new System.Windows.Forms.Label();
            lblTableName        = new System.Windows.Forms.Label();
            author              = new System.Windows.Forms.Label();
            label1              = new System.Windows.Forms.Label();
            gbType.SuspendLayout();
            SuspendLayout();

            // txtDimension
            txtDimension.Font     = new System.Drawing.Font("Segoe UI", 13F);
            txtDimension.Location = new System.Drawing.Point(60, 175);
            txtDimension.Name     = "txtDimension";
            txtDimension.Size     = new System.Drawing.Size(120, 31);
            txtDimension.TabIndex = 0;
            txtDimension.KeyPress    += TxtDimension_KeyPress;
            txtDimension.TextChanged += TxtDimension_TextChanged;

            // txtDimensionInch
            txtDimensionInch.Font     = new System.Drawing.Font("Segoe UI", 13F);
            txtDimensionInch.Location = new System.Drawing.Point(215, 175);
            txtDimensionInch.Name     = "txtDimensionInch";
            txtDimensionInch.Size     = new System.Drawing.Size(120, 31);
            txtDimensionInch.TabIndex = 1;
            txtDimensionInch.KeyPress    += TxtDimension_KeyPress;
            txtDimensionInch.TextChanged += TxtDimensionInch_TextChanged;

            // rbInternal — all three radio buttons on a single row
            rbInternal.AutoSize = true;
            rbInternal.Checked  = true;
            rbInternal.Location = new System.Drawing.Point(18, 24);
            rbInternal.Name     = "rbInternal";
            rbInternal.TabIndex = 0;
            rbInternal.TabStop  = true;
            rbInternal.Text     = "Internal (Hole)";
            rbInternal.UseVisualStyleBackColor = true;
            rbInternal.CheckedChanged += RbType_CheckedChanged;

            // rbExternal
            rbExternal.AutoSize = true;
            rbExternal.Location = new System.Drawing.Point(155, 24);
            rbExternal.Name     = "rbExternal";
            rbExternal.TabIndex = 1;
            rbExternal.Text     = "External (Shaft)";
            rbExternal.UseVisualStyleBackColor = true;
            rbExternal.CheckedChanged += RbType_CheckedChanged;

            // rbIT12Half
            rbIT12Half.AutoSize = true;
            rbIT12Half.Location = new System.Drawing.Point(296, 24);
            rbIT12Half.Name     = "rbIT12Half";
            rbIT12Half.TabIndex = 2;
            rbIT12Half.Text     = "IT12/2";
            rbIT12Half.UseVisualStyleBackColor = true;
            rbIT12Half.CheckedChanged += RbType_CheckedChanged;

            // btnConvert
            btnConvert.Font     = new System.Drawing.Font("Segoe UI", 10F);
            btnConvert.Location = new System.Drawing.Point(348, 174);
            btnConvert.Name     = "btnConvert";
            btnConvert.Size     = new System.Drawing.Size(83, 34);
            btnConvert.TabIndex = 2;
            btnConvert.Text     = "Convert";
            btnConvert.UseVisualStyleBackColor = true;
            btnConvert.Click += BtnConvert_Click;

            // lblUpperValue
            lblUpperValue.AutoSize = true;
            lblUpperValue.Font     = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold);
            lblUpperValue.Location = new System.Drawing.Point(150, 278);
            lblUpperValue.Name     = "lblUpperValue";
            lblUpperValue.TabIndex = 0;
            lblUpperValue.Text     = "-";

            // lblLowerValue
            lblLowerValue.AutoSize = true;
            lblLowerValue.Font     = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold);
            lblLowerValue.Location = new System.Drawing.Point(150, 338);
            lblLowerValue.Name     = "lblLowerValue";
            lblLowerValue.TabIndex = 0;
            lblLowerValue.Text     = "-";

            // lblError
            lblError.ForeColor = System.Drawing.Color.Red;
            lblError.Location  = new System.Drawing.Point(30, 152);
            lblError.Name      = "lblError";
            lblError.Size      = new System.Drawing.Size(415, 20);
            lblError.TabIndex  = 0;
            lblError.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;

            // lblTitle
            lblTitle.AutoSize = true;
            lblTitle.Font     = new System.Drawing.Font("Arial", 14F, System.Drawing.FontStyle.Bold);
            lblTitle.Location = new System.Drawing.Point(155, 20);
            lblTitle.Name     = "lblTitle";
            lblTitle.TabIndex = 0;
            lblTitle.Text     = "Tolerance Chart";

            // lblDimension
            lblDimension.AutoSize = true;
            lblDimension.Font     = new System.Drawing.Font("Segoe UI", 10F);
            lblDimension.Location = new System.Drawing.Point(26, 181);
            lblDimension.Name     = "lblDimension";
            lblDimension.TabIndex = 0;
            lblDimension.Text     = "mm:";

            // lblDimensionInch
            lblDimensionInch.AutoSize = true;
            lblDimensionInch.Font     = new System.Drawing.Font("Segoe UI", 10F);
            lblDimensionInch.Location = new System.Drawing.Point(190, 181);
            lblDimensionInch.Name     = "lblDimensionInch";
            lblDimensionInch.TabIndex = 0;
            lblDimensionInch.Text     = "in:";

            // gbType — single row, height reduced to 60
            gbType.Controls.Add(rbInternal);
            gbType.Controls.Add(rbExternal);
            gbType.Controls.Add(rbIT12Half);
            gbType.Font     = new System.Drawing.Font("Segoe UI", 10F);
            gbType.Location = new System.Drawing.Point(30, 79);
            gbType.Name     = "gbType";
            gbType.Size     = new System.Drawing.Size(415, 60);
            gbType.TabIndex = 3;
            gbType.TabStop  = false;
            gbType.Text     = "Type";

            // lblUpper
            lblUpper.AutoSize = true;
            lblUpper.Font     = new System.Drawing.Font("Segoe UI", 10F);
            lblUpper.Location = new System.Drawing.Point(150, 248);
            lblUpper.Name     = "lblUpper";
            lblUpper.TabIndex = 0;
            lblUpper.Text     = "Upper Tolerance:";

            // lblLower
            lblLower.AutoSize = true;
            lblLower.Font     = new System.Drawing.Font("Segoe UI", 10F);
            lblLower.Location = new System.Drawing.Point(150, 308);
            lblLower.Name     = "lblLower";
            lblLower.TabIndex = 0;
            lblLower.Text     = "Lower Tolerance:";

            // lblDescription
            lblDescription.AutoSize = true;
            lblDescription.Font     = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Italic);
            lblDescription.Location = new System.Drawing.Point(183, 54);
            lblDescription.Name     = "lblDescription";
            lblDescription.TabIndex = 0;
            lblDescription.Text     = "PER ASME B4.2-1978";

            // lblTableName
            lblTableName.AutoSize = true;
            lblTableName.Font     = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold);
            lblTableName.Location = new System.Drawing.Point(36, 260);
            lblTableName.Name     = "lblTableName";
            lblTableName.TabIndex = 0;
            lblTableName.Text     = "H12";

            // author
            author.AutoSize  = true;
            author.Font      = new System.Drawing.Font("Segoe UI", 8F);
            author.ForeColor = System.Drawing.SystemColors.GrayText;
            author.Location  = new System.Drawing.Point(116, 374);
            author.Name      = "author";
            author.TabIndex  = 4;
            author.Text      = "© 2026 Record Technology and Development";

            // label1
            label1.AutoSize  = true;
            label1.Font      = new System.Drawing.Font("Segoe UI", 8F);
            label1.ForeColor = System.Drawing.SystemColors.GrayText;
            label1.Location  = new System.Drawing.Point(130, 387);
            label1.Name      = "label1";
            label1.TabIndex  = 5;
            label1.Text      = "All Rights Reserved. Developed by Leon.";

            // ToleranceConverterForm
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode       = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize          = new System.Drawing.Size(470, 415);
            Controls.Add(label1);
            Controls.Add(author);
            Controls.Add(lblLower);
            Controls.Add(lblUpper);
            Controls.Add(gbType);
            Controls.Add(lblDimension);
            Controls.Add(lblDimensionInch);
            Controls.Add(lblTitle);
            Controls.Add(lblError);
            Controls.Add(lblLowerValue);
            Controls.Add(lblUpperValue);
            Controls.Add(btnConvert);
            Controls.Add(txtDimensionInch);
            Controls.Add(txtDimension);
            Controls.Add(lblDescription);
            Controls.Add(lblTableName);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            Name            = "ToleranceConverterForm";
            StartPosition   = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text            = "Tolerance Converter";
            gbType.ResumeLayout(false);
            gbType.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }
    }
}
