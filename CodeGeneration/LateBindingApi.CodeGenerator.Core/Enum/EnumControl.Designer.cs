namespace LateBindingApi.CodeGenerator.Core
{
    partial class EnumControl
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridViewComponents = new System.Windows.Forms.DataGridView();
            this.labelComponentsCaption = new System.Windows.Forms.Label();
            this.labelEnumCaption = new System.Windows.Forms.Label();
            this.dataGridViewEnum = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewComponents)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEnum)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridViewComponents
            // 
            this.dataGridViewComponents.AllowUserToAddRows = false;
            this.dataGridViewComponents.AllowUserToDeleteRows = false;
            this.dataGridViewComponents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewComponents.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewComponents.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewComponents.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridViewComponents.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewComponents.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridViewComponents.EnableHeadersVisualStyles = false;
            this.dataGridViewComponents.Location = new System.Drawing.Point(0, 22);
            this.dataGridViewComponents.MultiSelect = false;
            this.dataGridViewComponents.Name = "dataGridViewComponents";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewComponents.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridViewComponents.RowHeadersVisible = false;
            this.dataGridViewComponents.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewComponents.ShowCellToolTips = false;
            this.dataGridViewComponents.ShowEditingIcon = false;
            this.dataGridViewComponents.Size = new System.Drawing.Size(705, 98);
            this.dataGridViewComponents.TabIndex = 11;
            // 
            // labelComponentsCaption
            // 
            this.labelComponentsCaption.AutoSize = true;
            this.labelComponentsCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelComponentsCaption.Location = new System.Drawing.Point(0, 0);
            this.labelComponentsCaption.Name = "labelComponentsCaption";
            this.labelComponentsCaption.Size = new System.Drawing.Size(330, 20);
            this.labelComponentsCaption.TabIndex = 10;
            this.labelComponentsCaption.Text = "Enum is defined in the following Components.";
            // 
            // labelEnumCaption
            // 
            this.labelEnumCaption.AutoSize = true;
            this.labelEnumCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelEnumCaption.Location = new System.Drawing.Point(3, 138);
            this.labelEnumCaption.Name = "labelEnumCaption";
            this.labelEnumCaption.Size = new System.Drawing.Size(104, 20);
            this.labelEnumCaption.TabIndex = 9;
            this.labelEnumCaption.Text = "Enum Details";
            // 
            // dataGridViewEnum
            // 
            this.dataGridViewEnum.AllowUserToAddRows = false;
            this.dataGridViewEnum.AllowUserToDeleteRows = false;
            this.dataGridViewEnum.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewEnum.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewEnum.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewEnum.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewEnum.EnableHeadersVisualStyles = false;
            this.dataGridViewEnum.Location = new System.Drawing.Point(0, 161);
            this.dataGridViewEnum.MultiSelect = false;
            this.dataGridViewEnum.Name = "dataGridViewEnum";
            this.dataGridViewEnum.RowHeadersVisible = false;
            this.dataGridViewEnum.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewEnum.ShowCellToolTips = false;
            this.dataGridViewEnum.ShowEditingIcon = false;
            this.dataGridViewEnum.Size = new System.Drawing.Size(705, 227);
            this.dataGridViewEnum.TabIndex = 8;
            // 
            // EnumControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.dataGridViewComponents);
            this.Controls.Add(this.labelComponentsCaption);
            this.Controls.Add(this.labelEnumCaption);
            this.Controls.Add(this.dataGridViewEnum);
            this.Name = "EnumControl";
            this.Size = new System.Drawing.Size(705, 393);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewComponents)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEnum)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewComponents;
        private System.Windows.Forms.Label labelComponentsCaption;
        private System.Windows.Forms.Label labelEnumCaption;
        private System.Windows.Forms.DataGridView dataGridViewEnum;
    }
}
