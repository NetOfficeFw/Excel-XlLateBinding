namespace LateBindingApi.CodeGenerator
{
    partial class MethodsControl
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
            this.dataGridViewMethodParams = new System.Windows.Forms.DataGridView();
            this.dataGridViewMethods = new System.Windows.Forms.DataGridView();
            this.labelParametersCaption = new System.Windows.Forms.Label();
            this.labelFilterCaption = new System.Windows.Forms.Label();
            this.textBoxMethodFilter = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMethodParams)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMethods)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridViewMethodParams
            // 
            this.dataGridViewMethodParams.AllowUserToAddRows = false;
            this.dataGridViewMethodParams.AllowUserToDeleteRows = false;
            this.dataGridViewMethodParams.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewMethodParams.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewMethodParams.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewMethodParams.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewMethodParams.EnableHeadersVisualStyles = false;
            this.dataGridViewMethodParams.Location = new System.Drawing.Point(0, 241);
            this.dataGridViewMethodParams.MultiSelect = false;
            this.dataGridViewMethodParams.Name = "dataGridViewMethodParams";
            this.dataGridViewMethodParams.RowHeadersVisible = false;
            this.dataGridViewMethodParams.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewMethodParams.ShowCellToolTips = false;
            this.dataGridViewMethodParams.ShowEditingIcon = false;
            this.dataGridViewMethodParams.Size = new System.Drawing.Size(777, 158);
            this.dataGridViewMethodParams.TabIndex = 15;
            this.dataGridViewMethodParams.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewMethodParams_CellValueChanged);
            // 
            // dataGridViewMethods
            // 
            this.dataGridViewMethods.AllowUserToAddRows = false;
            this.dataGridViewMethods.AllowUserToDeleteRows = false;
            this.dataGridViewMethods.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewMethods.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewMethods.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewMethods.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewMethods.EnableHeadersVisualStyles = false;
            this.dataGridViewMethods.Location = new System.Drawing.Point(0, 38);
            this.dataGridViewMethods.MultiSelect = false;
            this.dataGridViewMethods.Name = "dataGridViewMethods";
            this.dataGridViewMethods.RowHeadersVisible = false;
            this.dataGridViewMethods.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewMethods.ShowCellToolTips = false;
            this.dataGridViewMethods.ShowEditingIcon = false;
            this.dataGridViewMethods.Size = new System.Drawing.Size(777, 167);
            this.dataGridViewMethods.TabIndex = 14;
            this.dataGridViewMethods.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewMethods_CellValueChanged);
            this.dataGridViewMethods.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewMethods_CellClick);
            this.dataGridViewMethods.SelectionChanged += new System.EventHandler(this.dataGridViewMethods_SelectionChanged);
            // 
            // labelParametersCaption
            // 
            this.labelParametersCaption.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelParametersCaption.AutoSize = true;
            this.labelParametersCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelParametersCaption.Location = new System.Drawing.Point(3, 218);
            this.labelParametersCaption.Name = "labelParametersCaption";
            this.labelParametersCaption.Size = new System.Drawing.Size(91, 20);
            this.labelParametersCaption.TabIndex = 16;
            this.labelParametersCaption.Text = "Parameters";
            // 
            // labelFilterCaption
            // 
            this.labelFilterCaption.AutoSize = true;
            this.labelFilterCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelFilterCaption.Location = new System.Drawing.Point(3, 14);
            this.labelFilterCaption.Name = "labelFilterCaption";
            this.labelFilterCaption.Size = new System.Drawing.Size(44, 20);
            this.labelFilterCaption.TabIndex = 17;
            this.labelFilterCaption.Text = "Filter";
            // 
            // textBoxMethodFilter
            // 
            this.textBoxMethodFilter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxMethodFilter.Location = new System.Drawing.Point(53, 11);
            this.textBoxMethodFilter.Name = "textBoxMethodFilter";
            this.textBoxMethodFilter.Size = new System.Drawing.Size(228, 22);
            this.textBoxMethodFilter.TabIndex = 18;
            this.textBoxMethodFilter.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBoxMethodFilter_KeyDown);
            // 
            // MethodsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.textBoxMethodFilter);
            this.Controls.Add(this.labelFilterCaption);
            this.Controls.Add(this.labelParametersCaption);
            this.Controls.Add(this.dataGridViewMethodParams);
            this.Controls.Add(this.dataGridViewMethods);
            this.Name = "MethodsControl";
            this.Size = new System.Drawing.Size(777, 399);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMethodParams)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMethods)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewMethodParams;
        private System.Windows.Forms.DataGridView dataGridViewMethods;
        private System.Windows.Forms.Label labelParametersCaption;
        private System.Windows.Forms.Label labelFilterCaption;
        private System.Windows.Forms.TextBox textBoxMethodFilter;
    }
}
