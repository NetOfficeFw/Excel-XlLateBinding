namespace LateBindingApi.CodeGenerator.Core.Visual
{
    partial class RefComponentsGrid
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
            this.dataGridViewComponents = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewComponents)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridViewComponents
            // 
            this.dataGridViewComponents.AllowUserToAddRows = false;
            this.dataGridViewComponents.AllowUserToDeleteRows = false;
            this.dataGridViewComponents.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewComponents.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewComponents.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewComponents.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewComponents.EnableHeadersVisualStyles = false;
            this.dataGridViewComponents.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewComponents.MultiSelect = false;
            this.dataGridViewComponents.Name = "dataGridViewComponents";
            this.dataGridViewComponents.RowHeadersVisible = false;
            this.dataGridViewComponents.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewComponents.ShowCellToolTips = false;
            this.dataGridViewComponents.ShowEditingIcon = false;
            this.dataGridViewComponents.Size = new System.Drawing.Size(688, 233);
            this.dataGridViewComponents.TabIndex = 10;
            // 
            // RefComponentsGrid
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.dataGridViewComponents);
            this.Name = "RefComponentsGrid";
            this.Size = new System.Drawing.Size(688, 233);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewComponents)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewComponents;
    }
}
