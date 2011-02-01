namespace LateBindingApi.CodeGenerator
{
    partial class ComponentsControl
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
            this.textBoxSolutionName = new System.Windows.Forms.TextBox();
            this.dataGridViewComponents = new System.Windows.Forms.DataGridView();
            this.labelSolutionName = new System.Windows.Forms.Label();
            this.textBoxClassPrefix = new System.Windows.Forms.TextBox();
            this.labelClassPrefix = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewComponents)).BeginInit();
            this.SuspendLayout();
            // 
            // textBoxSolutionName
            // 
            this.textBoxSolutionName.BackColor = System.Drawing.Color.DarkKhaki;
            this.textBoxSolutionName.Location = new System.Drawing.Point(97, 17);
            this.textBoxSolutionName.Name = "textBoxSolutionName";
            this.textBoxSolutionName.Size = new System.Drawing.Size(232, 20);
            this.textBoxSolutionName.TabIndex = 5;
            this.textBoxSolutionName.TextChanged += new System.EventHandler(this.textBoxSolutionName_TextChanged);
            // 
            // dataGridViewComponents
            // 
            this.dataGridViewComponents.AllowUserToAddRows = false;
            this.dataGridViewComponents.AllowUserToDeleteRows = false;
            this.dataGridViewComponents.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewComponents.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewComponents.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewComponents.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewComponents.EnableHeadersVisualStyles = false;
            this.dataGridViewComponents.Location = new System.Drawing.Point(0, 201);
            this.dataGridViewComponents.MultiSelect = false;
            this.dataGridViewComponents.Name = "dataGridViewComponents";
            this.dataGridViewComponents.RowHeadersVisible = false;
            this.dataGridViewComponents.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewComponents.ShowCellToolTips = false;
            this.dataGridViewComponents.ShowEditingIcon = false;
            this.dataGridViewComponents.Size = new System.Drawing.Size(705, 192);
            this.dataGridViewComponents.TabIndex = 3;
            this.dataGridViewComponents.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewComponents_CellValueChanged);
            // 
            // labelSolutionName
            // 
            this.labelSolutionName.AutoSize = true;
            this.labelSolutionName.Location = new System.Drawing.Point(12, 24);
            this.labelSolutionName.Name = "labelSolutionName";
            this.labelSolutionName.Size = new System.Drawing.Size(79, 13);
            this.labelSolutionName.TabIndex = 4;
            this.labelSolutionName.Text = "Solution Name:";
            // 
            // textBoxClassPrefix
            // 
            this.textBoxClassPrefix.BackColor = System.Drawing.Color.DarkKhaki;
            this.textBoxClassPrefix.Location = new System.Drawing.Point(97, 52);
            this.textBoxClassPrefix.Name = "textBoxClassPrefix";
            this.textBoxClassPrefix.Size = new System.Drawing.Size(50, 20);
            this.textBoxClassPrefix.TabIndex = 7;
            this.textBoxClassPrefix.TextChanged += new System.EventHandler(this.textBoxSolutionName_TextChanged);
            // 
            // labelClassPrefix
            // 
            this.labelClassPrefix.AutoSize = true;
            this.labelClassPrefix.Location = new System.Drawing.Point(12, 59);
            this.labelClassPrefix.Name = "labelClassPrefix";
            this.labelClassPrefix.Size = new System.Drawing.Size(64, 13);
            this.labelClassPrefix.TabIndex = 6;
            this.labelClassPrefix.Text = "Class Prefix:";
            // 
            // ComponentsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.textBoxClassPrefix);
            this.Controls.Add(this.labelClassPrefix);
            this.Controls.Add(this.textBoxSolutionName);
            this.Controls.Add(this.dataGridViewComponents);
            this.Controls.Add(this.labelSolutionName);
            this.Name = "ComponentsControl";
            this.Size = new System.Drawing.Size(705, 393);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewComponents)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxSolutionName;
        private System.Windows.Forms.DataGridView dataGridViewComponents;
        private System.Windows.Forms.Label labelSolutionName;
        private System.Windows.Forms.TextBox textBoxClassPrefix;
        private System.Windows.Forms.Label labelClassPrefix;
    }
}
