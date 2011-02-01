namespace LateBindingApi.CodeGenerator
{
    partial class PropertiesControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBoxMethodFilter = new System.Windows.Forms.TextBox();
            this.labelFilterCaption = new System.Windows.Forms.Label();
            this.labelParametersCaption = new System.Windows.Forms.Label();
            this.dataGridViewPropertiesParams = new System.Windows.Forms.DataGridView();
            this.dataGridViewProperties = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewPropertiesParams)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewProperties)).BeginInit();
            this.SuspendLayout();
            // 
            // textBoxMethodFilter
            // 
            this.textBoxMethodFilter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxMethodFilter.Location = new System.Drawing.Point(53, 5);
            this.textBoxMethodFilter.Name = "textBoxMethodFilter";
            this.textBoxMethodFilter.Size = new System.Drawing.Size(228, 22);
            this.textBoxMethodFilter.TabIndex = 23;
            this.textBoxMethodFilter.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBoxMethodFilter_KeyDown);
            // 
            // labelFilterCaption
            // 
            this.labelFilterCaption.AutoSize = true;
            this.labelFilterCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelFilterCaption.Location = new System.Drawing.Point(3, 8);
            this.labelFilterCaption.Name = "labelFilterCaption";
            this.labelFilterCaption.Size = new System.Drawing.Size(44, 20);
            this.labelFilterCaption.TabIndex = 22;
            this.labelFilterCaption.Text = "Filter";
            // 
            // labelParametersCaption
            // 
            this.labelParametersCaption.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelParametersCaption.AutoSize = true;
            this.labelParametersCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelParametersCaption.Location = new System.Drawing.Point(3, 212);
            this.labelParametersCaption.Name = "labelParametersCaption";
            this.labelParametersCaption.Size = new System.Drawing.Size(91, 20);
            this.labelParametersCaption.TabIndex = 21;
            this.labelParametersCaption.Text = "Parameters";
            // 
            // dataGridViewPropertiesParams
            // 
            this.dataGridViewPropertiesParams.AllowUserToAddRows = false;
            this.dataGridViewPropertiesParams.AllowUserToDeleteRows = false;
            this.dataGridViewPropertiesParams.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewPropertiesParams.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewPropertiesParams.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewPropertiesParams.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewPropertiesParams.EnableHeadersVisualStyles = false;
            this.dataGridViewPropertiesParams.Location = new System.Drawing.Point(0, 235);
            this.dataGridViewPropertiesParams.MultiSelect = false;
            this.dataGridViewPropertiesParams.Name = "dataGridViewPropertiesParams";
            this.dataGridViewPropertiesParams.RowHeadersVisible = false;
            this.dataGridViewPropertiesParams.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewPropertiesParams.ShowCellToolTips = false;
            this.dataGridViewPropertiesParams.ShowEditingIcon = false;
            this.dataGridViewPropertiesParams.Size = new System.Drawing.Size(777, 158);
            this.dataGridViewPropertiesParams.TabIndex = 20;
            this.dataGridViewPropertiesParams.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewMethodParams_CellValueChanged);
            // 
            // dataGridViewProperties
            // 
            this.dataGridViewProperties.AllowUserToAddRows = false;
            this.dataGridViewProperties.AllowUserToDeleteRows = false;
            this.dataGridViewProperties.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewProperties.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewProperties.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewProperties.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewProperties.EnableHeadersVisualStyles = false;
            this.dataGridViewProperties.Location = new System.Drawing.Point(0, 32);
            this.dataGridViewProperties.MultiSelect = false;
            this.dataGridViewProperties.Name = "dataGridViewProperties";
            this.dataGridViewProperties.RowHeadersVisible = false;
            this.dataGridViewProperties.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewProperties.ShowCellToolTips = false;
            this.dataGridViewProperties.ShowEditingIcon = false;
            this.dataGridViewProperties.Size = new System.Drawing.Size(777, 167);
            this.dataGridViewProperties.TabIndex = 19;
            this.dataGridViewProperties.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewMethods_CellValueChanged);
            this.dataGridViewProperties.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewMethods_CellClick);
            this.dataGridViewProperties.SelectionChanged += new System.EventHandler(this.dataGridViewMethods_SelectionChanged);
            // 
            // PropertiesControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.textBoxMethodFilter);
            this.Controls.Add(this.labelFilterCaption);
            this.Controls.Add(this.labelParametersCaption);
            this.Controls.Add(this.dataGridViewPropertiesParams);
            this.Controls.Add(this.dataGridViewProperties);
            this.Name = "PropertiesControl";
            this.Size = new System.Drawing.Size(777, 399);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewPropertiesParams)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewProperties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxMethodFilter;
        private System.Windows.Forms.Label labelFilterCaption;
        private System.Windows.Forms.Label labelParametersCaption;
        private System.Windows.Forms.DataGridView dataGridViewPropertiesParams;
        private System.Windows.Forms.DataGridView dataGridViewProperties;
    }
}
