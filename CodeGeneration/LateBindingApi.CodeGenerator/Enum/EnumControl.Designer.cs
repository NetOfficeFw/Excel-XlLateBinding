namespace LateBindingApi.CodeGenerator
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle12 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridViewEnum = new System.Windows.Forms.DataGridView();
            this.labelEnumCaption = new System.Windows.Forms.Label();
            this.labelComponentsCaption = new System.Windows.Forms.Label();
            this.dataGridViewComponents = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEnum)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewComponents)).BeginInit();
            this.SuspendLayout();
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
            this.dataGridViewEnum.Location = new System.Drawing.Point(0, 197);
            this.dataGridViewEnum.MultiSelect = false;
            this.dataGridViewEnum.Name = "dataGridViewEnum";
            this.dataGridViewEnum.RowHeadersVisible = false;
            this.dataGridViewEnum.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewEnum.ShowCellToolTips = false;
            this.dataGridViewEnum.ShowEditingIcon = false;
            this.dataGridViewEnum.Size = new System.Drawing.Size(705, 196);
            this.dataGridViewEnum.TabIndex = 4;
            this.dataGridViewEnum.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewEnum_CellValueChanged);
            // 
            // labelEnumCaption
            // 
            this.labelEnumCaption.AutoSize = true;
            this.labelEnumCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelEnumCaption.Location = new System.Drawing.Point(3, 174);
            this.labelEnumCaption.Name = "labelEnumCaption";
            this.labelEnumCaption.Size = new System.Drawing.Size(104, 20);
            this.labelEnumCaption.TabIndex = 5;
            this.labelEnumCaption.Text = "Enum Details";
            // 
            // labelComponentsCaption
            // 
            this.labelComponentsCaption.AutoSize = true;
            this.labelComponentsCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelComponentsCaption.Location = new System.Drawing.Point(3, 10);
            this.labelComponentsCaption.Name = "labelComponentsCaption";
            this.labelComponentsCaption.Size = new System.Drawing.Size(330, 20);
            this.labelComponentsCaption.TabIndex = 6;
            this.labelComponentsCaption.Text = "Enum is defined in the following Components.";
            // 
            // dataGridViewComponents
            // 
            this.dataGridViewComponents.AllowUserToAddRows = false;
            this.dataGridViewComponents.AllowUserToDeleteRows = false;
            this.dataGridViewComponents.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewComponents.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewComponents.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            dataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewComponents.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle10;
            this.dataGridViewComponents.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle11.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewComponents.DefaultCellStyle = dataGridViewCellStyle11;
            this.dataGridViewComponents.EnableHeadersVisualStyles = false;
            this.dataGridViewComponents.Location = new System.Drawing.Point(0, 33);
            this.dataGridViewComponents.MultiSelect = false;
            this.dataGridViewComponents.Name = "dataGridViewComponents";
            dataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle12.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle12.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle12.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle12.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle12.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewComponents.RowHeadersDefaultCellStyle = dataGridViewCellStyle12;
            this.dataGridViewComponents.RowHeadersVisible = false;
            this.dataGridViewComponents.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewComponents.ShowCellToolTips = false;
            this.dataGridViewComponents.ShowEditingIcon = false;
            this.dataGridViewComponents.Size = new System.Drawing.Size(705, 91);
            this.dataGridViewComponents.TabIndex = 7;
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
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEnum)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewComponents)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewEnum;
        private System.Windows.Forms.Label labelEnumCaption;
        private System.Windows.Forms.Label labelComponentsCaption;
        private System.Windows.Forms.DataGridView dataGridViewComponents;
    }
}
