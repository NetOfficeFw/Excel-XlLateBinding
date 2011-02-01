namespace LateBindingApi.CodeGenerator.Core.Visual
{
    partial class EnumGrid
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
            this.dataGridViewEnumMembers = new System.Windows.Forms.DataGridView();
            this.dataGridViewRefComponents = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEnumMembers)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewRefComponents)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridViewEnumMembers
            // 
            this.dataGridViewEnumMembers.AllowUserToAddRows = false;
            this.dataGridViewEnumMembers.AllowUserToDeleteRows = false;
            this.dataGridViewEnumMembers.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewEnumMembers.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewEnumMembers.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewEnumMembers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewEnumMembers.EnableHeadersVisualStyles = false;
            this.dataGridViewEnumMembers.Location = new System.Drawing.Point(0, 95);
            this.dataGridViewEnumMembers.MultiSelect = false;
            this.dataGridViewEnumMembers.Name = "dataGridViewEnumMembers";
            this.dataGridViewEnumMembers.RowHeadersVisible = false;
            this.dataGridViewEnumMembers.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewEnumMembers.ShowCellToolTips = false;
            this.dataGridViewEnumMembers.ShowEditingIcon = false;
            this.dataGridViewEnumMembers.Size = new System.Drawing.Size(487, 234);
            this.dataGridViewEnumMembers.TabIndex = 9;
            this.dataGridViewEnumMembers.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewEnumMembers_CellValueChanged);
            // 
            // dataGridViewRefComponents
            // 
            this.dataGridViewRefComponents.AllowUserToAddRows = false;
            this.dataGridViewRefComponents.AllowUserToDeleteRows = false;
            this.dataGridViewRefComponents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewRefComponents.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewRefComponents.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewRefComponents.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewRefComponents.EnableHeadersVisualStyles = false;
            this.dataGridViewRefComponents.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewRefComponents.MultiSelect = false;
            this.dataGridViewRefComponents.Name = "dataGridViewRefComponents";
            this.dataGridViewRefComponents.RowHeadersVisible = false;
            this.dataGridViewRefComponents.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewRefComponents.ShowCellToolTips = false;
            this.dataGridViewRefComponents.ShowEditingIcon = false;
            this.dataGridViewRefComponents.Size = new System.Drawing.Size(487, 89);
            this.dataGridViewRefComponents.TabIndex = 10;
            // 
            // EnumGrid
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.dataGridViewRefComponents);
            this.Controls.Add(this.dataGridViewEnumMembers);
            this.Name = "EnumGrid";
            this.Size = new System.Drawing.Size(487, 329);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEnumMembers)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewRefComponents)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewEnumMembers;
        private System.Windows.Forms.DataGridView dataGridViewRefComponents;
    }
}
