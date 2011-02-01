namespace LateBindingApi.CodeGenerator.Core.Visual
{
    partial class FormResolveMethodConflict
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridViewMethods = new System.Windows.Forms.DataGridView();
            this.buttonOk = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.dataGridViewMethodParameters = new System.Windows.Forms.DataGridView();
            this.buttonResolveAll = new System.Windows.Forms.Button();
            this.buttonUnResolveAll = new System.Windows.Forms.Button();
            this.radioButtonAll = new System.Windows.Forms.RadioButton();
            this.radioButtonConflicts = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMethods)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMethodParameters)).BeginInit();
            this.SuspendLayout();
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
            this.dataGridViewMethods.Location = new System.Drawing.Point(-1, 0);
            this.dataGridViewMethods.MultiSelect = false;
            this.dataGridViewMethods.Name = "dataGridViewMethods";
            this.dataGridViewMethods.RowHeadersVisible = false;
            this.dataGridViewMethods.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewMethods.ShowCellToolTips = false;
            this.dataGridViewMethods.ShowEditingIcon = false;
            this.dataGridViewMethods.Size = new System.Drawing.Size(712, 343);
            this.dataGridViewMethods.TabIndex = 16;
            this.dataGridViewMethods.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dataGridViewMethods_CellValidating);
            this.dataGridViewMethods.SelectionChanged += new System.EventHandler(this.dataGridViewMethods_SelectionChanged);
            // 
            // buttonOk
            // 
            this.buttonOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOk.Location = new System.Drawing.Point(523, 561);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(74, 21);
            this.buttonOk.TabIndex = 18;
            this.buttonOk.Text = "Ok";
            this.buttonOk.UseVisualStyleBackColor = true;
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCancel.Location = new System.Drawing.Point(612, 561);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(74, 21);
            this.buttonCancel.TabIndex = 17;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // dataGridViewMethodParameters
            // 
            this.dataGridViewMethodParameters.AllowUserToAddRows = false;
            this.dataGridViewMethodParameters.AllowUserToDeleteRows = false;
            this.dataGridViewMethodParameters.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewMethodParameters.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewMethodParameters.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewMethodParameters.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewMethodParameters.EnableHeadersVisualStyles = false;
            this.dataGridViewMethodParameters.Location = new System.Drawing.Point(-1, 376);
            this.dataGridViewMethodParameters.MultiSelect = false;
            this.dataGridViewMethodParameters.Name = "dataGridViewMethodParameters";
            this.dataGridViewMethodParameters.RowHeadersVisible = false;
            this.dataGridViewMethodParameters.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewMethodParameters.ShowCellToolTips = false;
            this.dataGridViewMethodParameters.ShowEditingIcon = false;
            this.dataGridViewMethodParameters.Size = new System.Drawing.Size(712, 160);
            this.dataGridViewMethodParameters.TabIndex = 19;
            // 
            // buttonResolveAll
            // 
            this.buttonResolveAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonResolveAll.Location = new System.Drawing.Point(178, 349);
            this.buttonResolveAll.Name = "buttonResolveAll";
            this.buttonResolveAll.Size = new System.Drawing.Size(106, 21);
            this.buttonResolveAll.TabIndex = 20;
            this.buttonResolveAll.Text = "Check All";
            this.buttonResolveAll.UseVisualStyleBackColor = true;
            this.buttonResolveAll.Click += new System.EventHandler(this.buttonResolve_Click);
            // 
            // buttonUnResolveAll
            // 
            this.buttonUnResolveAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonUnResolveAll.Location = new System.Drawing.Point(303, 349);
            this.buttonUnResolveAll.Name = "buttonUnResolveAll";
            this.buttonUnResolveAll.Size = new System.Drawing.Size(106, 21);
            this.buttonUnResolveAll.TabIndex = 21;
            this.buttonUnResolveAll.Text = "Uncheck All";
            this.buttonUnResolveAll.UseVisualStyleBackColor = true;
            this.buttonUnResolveAll.Click += new System.EventHandler(this.buttonResolve_Click);
            // 
            // radioButtonAll
            // 
            this.radioButtonAll.AutoSize = true;
            this.radioButtonAll.Location = new System.Drawing.Point(29, 353);
            this.radioButtonAll.Name = "radioButtonAll";
            this.radioButtonAll.Size = new System.Drawing.Size(36, 17);
            this.radioButtonAll.TabIndex = 22;
            this.radioButtonAll.Text = "All";
            this.radioButtonAll.UseVisualStyleBackColor = true;
            this.radioButtonAll.CheckedChanged += new System.EventHandler(this.radioButton_CheckedChanged);
            // 
            // radioButtonConflicts
            // 
            this.radioButtonConflicts.AutoSize = true;
            this.radioButtonConflicts.Checked = true;
            this.radioButtonConflicts.Location = new System.Drawing.Point(88, 353);
            this.radioButtonConflicts.Name = "radioButtonConflicts";
            this.radioButtonConflicts.Size = new System.Drawing.Size(65, 17);
            this.radioButtonConflicts.TabIndex = 23;
            this.radioButtonConflicts.TabStop = true;
            this.radioButtonConflicts.Text = "Conflicts";
            this.radioButtonConflicts.UseVisualStyleBackColor = true;
            // 
            // FormResolveMethodConflict
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(710, 594);
            this.Controls.Add(this.radioButtonConflicts);
            this.Controls.Add(this.radioButtonAll);
            this.Controls.Add(this.buttonUnResolveAll);
            this.Controls.Add(this.buttonResolveAll);
            this.Controls.Add(this.dataGridViewMethodParameters);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.dataGridViewMethods);
            this.Name = "FormResolveMethodConflict";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Resolve Conflict Assistant ";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMethods)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMethodParameters)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewMethods;
        private System.Windows.Forms.Button buttonOk;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.DataGridView dataGridViewMethodParameters;
        private System.Windows.Forms.Button buttonResolveAll;
        private System.Windows.Forms.Button buttonUnResolveAll;
        private System.Windows.Forms.RadioButton radioButtonAll;
        private System.Windows.Forms.RadioButton radioButtonConflicts;
    }
}