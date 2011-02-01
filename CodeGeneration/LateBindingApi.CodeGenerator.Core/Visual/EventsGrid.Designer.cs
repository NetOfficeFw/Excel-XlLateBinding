namespace LateBindingApi.CodeGenerator.Core.Visual
{
    partial class EventsGrid
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
            this.textBoxInterfaceName = new System.Windows.Forms.TextBox();
            this.labelFilterCaption = new System.Windows.Forms.Label();
            this.dataGridViewMethods = new System.Windows.Forms.DataGridView();
            this.textBoxMethodFilter = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonResolve = new System.Windows.Forms.Button();
            this.parametersGridMain = new LateBindingApi.CodeGenerator.Core.Visual.ParametersGrid();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMethods)).BeginInit();
            this.SuspendLayout();
            // 
            // textBoxInterfaceName
            // 
            this.textBoxInterfaceName.BackColor = System.Drawing.SystemColors.Control;
            this.textBoxInterfaceName.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBoxInterfaceName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxInterfaceName.Location = new System.Drawing.Point(363, 19);
            this.textBoxInterfaceName.Name = "textBoxInterfaceName";
            this.textBoxInterfaceName.ReadOnly = true;
            this.textBoxInterfaceName.Size = new System.Drawing.Size(131, 15);
            this.textBoxInterfaceName.TabIndex = 23;
            // 
            // labelFilterCaption
            // 
            this.labelFilterCaption.AutoSize = true;
            this.labelFilterCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelFilterCaption.Location = new System.Drawing.Point(289, 16);
            this.labelFilterCaption.Name = "labelFilterCaption";
            this.labelFilterCaption.Size = new System.Drawing.Size(73, 20);
            this.labelFilterCaption.TabIndex = 22;
            this.labelFilterCaption.Text = "Interface";
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
            this.dataGridViewMethods.Location = new System.Drawing.Point(0, 44);
            this.dataGridViewMethods.MultiSelect = false;
            this.dataGridViewMethods.Name = "dataGridViewMethods";
            this.dataGridViewMethods.RowHeadersVisible = false;
            this.dataGridViewMethods.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewMethods.ShowCellToolTips = false;
            this.dataGridViewMethods.ShowEditingIcon = false;
            this.dataGridViewMethods.Size = new System.Drawing.Size(660, 238);
            this.dataGridViewMethods.TabIndex = 21;
            this.dataGridViewMethods.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewMethods_CellValueChanged);
            this.dataGridViewMethods.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewMethods_CellClick);
            this.dataGridViewMethods.SelectionChanged += new System.EventHandler(this.dataGridViewMethods_SelectionChanged);
            // 
            // textBoxMethodFilter
            // 
            this.textBoxMethodFilter.BackColor = System.Drawing.Color.DarkKhaki;
            this.textBoxMethodFilter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxMethodFilter.Location = new System.Drawing.Point(55, 16);
            this.textBoxMethodFilter.Name = "textBoxMethodFilter";
            this.textBoxMethodFilter.Size = new System.Drawing.Size(228, 22);
            this.textBoxMethodFilter.TabIndex = 26;
            this.textBoxMethodFilter.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBoxMethodFilter_KeyDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(5, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(44, 20);
            this.label1.TabIndex = 25;
            this.label1.Text = "Filter";
            // 
            // buttonResolve
            // 
            this.buttonResolve.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonResolve.Location = new System.Drawing.Point(499, 14);
            this.buttonResolve.Name = "buttonResolve";
            this.buttonResolve.Size = new System.Drawing.Size(131, 24);
            this.buttonResolve.TabIndex = 27;
            this.buttonResolve.Text = "Resolve Conflict";
            this.buttonResolve.UseVisualStyleBackColor = true;
            this.buttonResolve.Click += new System.EventHandler(this.buttonResolve_Click);
            // 
            // parametersGridMain
            // 
            this.parametersGridMain.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.parametersGridMain.Location = new System.Drawing.Point(0, 288);
            this.parametersGridMain.Name = "parametersGridMain";
            this.parametersGridMain.Size = new System.Drawing.Size(660, 152);
            this.parametersGridMain.TabIndex = 24;
            // 
            // EventsGrid
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.buttonResolve);
            this.Controls.Add(this.textBoxMethodFilter);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.parametersGridMain);
            this.Controls.Add(this.textBoxInterfaceName);
            this.Controls.Add(this.labelFilterCaption);
            this.Controls.Add(this.dataGridViewMethods);
            this.Name = "EventsGrid";
            this.Size = new System.Drawing.Size(660, 440);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMethods)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxInterfaceName;
        private System.Windows.Forms.Label labelFilterCaption;
        private System.Windows.Forms.DataGridView dataGridViewMethods;
        private ParametersGrid parametersGridMain;
        private System.Windows.Forms.TextBox textBoxMethodFilter;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonResolve;
    }
}
