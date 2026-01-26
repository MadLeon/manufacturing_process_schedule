namespace ToleranceConverter
{
    partial class ToleranceConverterForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TextBox txtDimension;
        private System.Windows.Forms.RadioButton rbInternal;
        private System.Windows.Forms.RadioButton rbExternal;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Label lblUpperValue;
        private System.Windows.Forms.Label lblLowerValue;
        private System.Windows.Forms.Label lblError;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblDimension;
        private System.Windows.Forms.GroupBox gbType;
        private System.Windows.Forms.Label lblUpper;
        private System.Windows.Forms.Label lblLower;
        private System.Windows.Forms.Label lblDescription;
        private System.Windows.Forms.Label lblTableName;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            txtDimension = new System.Windows.Forms.TextBox();
            rbInternal = new System.Windows.Forms.RadioButton();
            rbExternal = new System.Windows.Forms.RadioButton();
            btnConvert = new System.Windows.Forms.Button();
            lblUpperValue = new System.Windows.Forms.Label();
            lblLowerValue = new System.Windows.Forms.Label();
            lblError = new System.Windows.Forms.Label();
            lblTitle = new System.Windows.Forms.Label();
            lblDimension = new System.Windows.Forms.Label();
            gbType = new System.Windows.Forms.GroupBox();
            lblUpper = new System.Windows.Forms.Label();
            lblLower = new System.Windows.Forms.Label();
            author = new System.Windows.Forms.Label();
            label1 = new System.Windows.Forms.Label();
            lblDescription = new System.Windows.Forms.Label();
            lblTableName = new System.Windows.Forms.Label();
            gbType.SuspendLayout();
            SuspendLayout();
            // 
            // txtDimension
            // 
            txtDimension.Font = new System.Drawing.Font("Segoe UI", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            txtDimension.Location = new System.Drawing.Point(146, 163);
            txtDimension.Name = "txtDimension";
            txtDimension.Size = new System.Drawing.Size(157, 31);
            txtDimension.TabIndex = 0;
            txtDimension.KeyPress += TxtDimension_KeyPress;
            // 
            // rbInternal
            // 
            rbInternal.AutoSize = true;
            rbInternal.Checked = true;
            rbInternal.Location = new System.Drawing.Point(44, 24);
            rbInternal.Name = "rbInternal";
            rbInternal.Size = new System.Drawing.Size(114, 23);
            rbInternal.TabIndex = 0;
            rbInternal.TabStop = true;
            rbInternal.Text = "Internal (Hole)";
            rbInternal.UseVisualStyleBackColor = true;
            // 
            // rbExternal
            // 
            rbExternal.AutoSize = true;
            rbExternal.Location = new System.Drawing.Point(210, 24);
            rbExternal.Name = "rbExternal";
            rbExternal.Size = new System.Drawing.Size(118, 23);
            rbExternal.TabIndex = 1;
            rbExternal.Text = "External (Shaft)";
            rbExternal.UseVisualStyleBackColor = true;
            // 
            // btnConvert
            // 
            btnConvert.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            btnConvert.Location = new System.Drawing.Point(318, 162);
            btnConvert.Name = "btnConvert";
            btnConvert.Size = new System.Drawing.Size(83, 34);
            btnConvert.TabIndex = 1;
            btnConvert.Text = "Convert";
            btnConvert.UseVisualStyleBackColor = true;
            btnConvert.Click += BtnConvert_Click;
            // 
            // lblUpperValue
            // 
            lblUpperValue.AutoSize = true;
            lblUpperValue.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            lblUpperValue.Location = new System.Drawing.Point(260, 228);
            lblUpperValue.Name = "lblUpperValue";
            lblUpperValue.Size = new System.Drawing.Size(14, 19);
            lblUpperValue.TabIndex = 0;
            lblUpperValue.Text = "-";
            // 
            // lblLowerValue
            // 
            lblLowerValue.AutoSize = true;
            lblLowerValue.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            lblLowerValue.Location = new System.Drawing.Point(260, 258);
            lblLowerValue.Name = "lblLowerValue";
            lblLowerValue.Size = new System.Drawing.Size(14, 19);
            lblLowerValue.TabIndex = 0;
            lblLowerValue.Text = "-";
            // 
            // lblError
            // 
            lblError.ForeColor = System.Drawing.Color.Red;
            lblError.Location = new System.Drawing.Point(30, 124);
            lblError.Name = "lblError";
            lblError.Size = new System.Drawing.Size(370, 20);
            lblError.TabIndex = 0;
            lblError.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblTitle
            // 
            lblTitle.AutoSize = true;
            lblTitle.Font = new System.Drawing.Font("Arial", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            lblTitle.Location = new System.Drawing.Point(140, 20);
            lblTitle.Name = "lblTitle";
            lblTitle.Size = new System.Drawing.Size(159, 22);
            lblTitle.TabIndex = 0;
            lblTitle.Text = "Tolerance Chart";
            // 
            // lblDimension
            // 
            lblDimension.AutoSize = true;
            lblDimension.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            lblDimension.Location = new System.Drawing.Point(26, 169);
            lblDimension.Name = "lblDimension";
            lblDimension.Size = new System.Drawing.Size(113, 19);
            lblDimension.TabIndex = 0;
            lblDimension.Text = "Dimension (mm):";
            // 
            // gbType
            // 
            gbType.Controls.Add(rbInternal);
            gbType.Controls.Add(rbExternal);
            gbType.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            gbType.Location = new System.Drawing.Point(30, 79);
            gbType.Name = "gbType";
            gbType.Size = new System.Drawing.Size(370, 60);
            gbType.TabIndex = 2;
            gbType.TabStop = false;
            gbType.Text = "Type";
            // 
            // lblUpper
            // 
            lblUpper.AutoSize = true;
            lblUpper.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            lblUpper.Location = new System.Drawing.Point(140, 228);
            lblUpper.Name = "lblUpper";
            lblUpper.Size = new System.Drawing.Size(111, 19);
            lblUpper.TabIndex = 0;
            lblUpper.Text = "Upper Tolerance:";
            // 
            // lblLower
            // 
            lblLower.AutoSize = true;
            lblLower.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            lblLower.Location = new System.Drawing.Point(140, 258);
            lblLower.Name = "lblLower";
            lblLower.Size = new System.Drawing.Size(110, 19);
            lblLower.TabIndex = 0;
            lblLower.Text = "Lower Tolerance:";
            // 
            // author
            // 
            author.AutoSize = true;
            author.BackColor = System.Drawing.SystemColors.Control;
            author.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            author.ForeColor = System.Drawing.SystemColors.GrayText;
            author.Location = new System.Drawing.Point(105, 306);
            author.Name = "author";
            author.Size = new System.Drawing.Size(238, 13);
            author.TabIndex = 3;
            author.Text = "© 2026 Record Technology and Development";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.BackColor = System.Drawing.SystemColors.Control;
            label1.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            label1.ForeColor = System.Drawing.SystemColors.GrayText;
            label1.Location = new System.Drawing.Point(119, 319);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(212, 13);
            label1.TabIndex = 4;
            label1.Text = "All Rights Reserved. Developed by Leon.";
            // 
            // lblDescription
            // 
            lblDescription.AutoSize = true;
            lblDescription.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point);
            lblDescription.ForeColor = System.Drawing.Color.Black;
            lblDescription.Location = new System.Drawing.Point(167, 54);
            lblDescription.Name = "lblDescription";
            lblDescription.Size = new System.Drawing.Size(104, 13);
            lblDescription.TabIndex = 0;
            lblDescription.Text = "PER ASME B4.2-1978";
            // 
            // lblTableName
            // 
            lblTableName.AutoSize = true;
            lblTableName.Font = new System.Drawing.Font("Segoe UI", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            lblTableName.Location = new System.Drawing.Point(36, 228);
            lblTableName.Name = "lblTableName";
            lblTableName.Size = new System.Drawing.Size(81, 45);
            lblTableName.TabIndex = 0;
            lblTableName.Text = "H12";
            // 
            // ToleranceConverterForm
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(434, 359);
            Controls.Add(label1);
            Controls.Add(author);
            Controls.Add(lblLower);
            Controls.Add(lblUpper);
            Controls.Add(gbType);
            Controls.Add(lblDimension);
            Controls.Add(lblTitle);
            Controls.Add(lblError);
            Controls.Add(lblLowerValue);
            Controls.Add(lblUpperValue);
            Controls.Add(btnConvert);
            Controls.Add(txtDimension);
            Controls.Add(lblDescription);
            Controls.Add(lblTableName);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            Name = "ToleranceConverterForm";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Tolerance Converter";
            gbType.ResumeLayout(false);
            gbType.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        private System.Windows.Forms.Label author;
        private System.Windows.Forms.Label label1;
    }
}
