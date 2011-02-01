namespace LateBindingApi.CodeGenerator.Core.Visual
{
    partial class RefInterfacesGrid
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
            this.dataGridViewRefInterfaces = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewRefInterfaces)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridViewRefInterfaces
            // 
            this.dataGridViewRefInterfaces.AllowUserToAddRows = false;
            this.dataGridViewRefInterfaces.AllowUserToDeleteRows = false;
            this.dataGridViewRefInterfaces.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewRefInterfaces.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewRefInterfaces.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewRefInterfaces.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewRefInterfaces.EnableHeadersVisualStyles = false;
            this.dataGridViewRefInterfaces.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewRefInterfaces.MultiSelect = false;
            this.dataGridViewRefInterfaces.Name = "dataGridViewRefInterfaces";
            this.dataGridViewRefInterfaces.RowHeadersVisible = false;
            this.dataGridViewRefInterfaces.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewRefInterfaces.ShowCellToolTips = false;
            this.dataGridViewRefInterfaces.ShowEditingIcon = false;
            this.dataGridViewRefInterfaces.Size = new System.Drawing.Size(456, 261);
            this.dataGridViewRefInterfaces.TabIndex = 11;
            // 
            // RefInterfacesGrid
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.dataGridViewRefInterfaces);
            this.Name = "RefInterfacesGrid";
            this.Size = new System.Drawing.Size(456, 261);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewRefInterfaces)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewRefInterfaces;
    }
}
