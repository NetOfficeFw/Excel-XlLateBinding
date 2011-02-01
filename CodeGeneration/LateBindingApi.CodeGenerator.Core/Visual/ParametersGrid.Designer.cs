namespace LateBindingApi.CodeGenerator.Core.Visual
{
    partial class ParametersGrid
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
            this.dataGridViewMethodParameters = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMethodParameters)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridViewMethodParameters
            // 
            this.dataGridViewMethodParameters.AllowUserToAddRows = false;
            this.dataGridViewMethodParameters.AllowUserToDeleteRows = false;
            this.dataGridViewMethodParameters.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewMethodParameters.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewMethodParameters.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewMethodParameters.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewMethodParameters.EnableHeadersVisualStyles = false;
            this.dataGridViewMethodParameters.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewMethodParameters.MultiSelect = false;
            this.dataGridViewMethodParameters.Name = "dataGridViewMethodParameters";
            this.dataGridViewMethodParameters.RowHeadersVisible = false;
            this.dataGridViewMethodParameters.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewMethodParameters.ShowCellToolTips = false;
            this.dataGridViewMethodParameters.ShowEditingIcon = false;
            this.dataGridViewMethodParameters.Size = new System.Drawing.Size(479, 170);
            this.dataGridViewMethodParameters.TabIndex = 17;
            this.dataGridViewMethodParameters.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewMethodParameters_CellValueChanged);
            this.dataGridViewMethodParameters.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewMethodParameters_CellClick);
            // 
            // ParametersGrid
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.dataGridViewMethodParameters);
            this.Name = "ParametersGrid";
            this.Size = new System.Drawing.Size(479, 170);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMethodParameters)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewMethodParameters;
    }
}
