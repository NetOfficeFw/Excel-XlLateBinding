namespace LateBindingApi.CodeGenerator
{
    partial class EventsControl
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
            this.textBoxMethodFilter = new System.Windows.Forms.TextBox();
            this.labelFilterCaption = new System.Windows.Forms.Label();
            this.labelParametersCaption = new System.Windows.Forms.Label();
            this.dataGridViewEventsParams = new System.Windows.Forms.DataGridView();
            this.dataGridViewEvents = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEventsParams)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEvents)).BeginInit();
            this.SuspendLayout();
            // 
            // textBoxMethodFilter
            // 
            this.textBoxMethodFilter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxMethodFilter.Location = new System.Drawing.Point(53, 5);
            this.textBoxMethodFilter.Name = "textBoxMethodFilter";
            this.textBoxMethodFilter.Size = new System.Drawing.Size(228, 22);
            this.textBoxMethodFilter.TabIndex = 23;
            this.textBoxMethodFilter.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBoxMethodFilter_KeyDown);
            // 
            // labelFilterCaption
            // 
            this.labelFilterCaption.AutoSize = true;
            this.labelFilterCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelFilterCaption.Location = new System.Drawing.Point(3, 8);
            this.labelFilterCaption.Name = "labelFilterCaption";
            this.labelFilterCaption.Size = new System.Drawing.Size(44, 20);
            this.labelFilterCaption.TabIndex = 22;
            this.labelFilterCaption.Text = "Filter";
            // 
            // labelParametersCaption
            // 
            this.labelParametersCaption.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelParametersCaption.AutoSize = true;
            this.labelParametersCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelParametersCaption.Location = new System.Drawing.Point(3, 212);
            this.labelParametersCaption.Name = "labelParametersCaption";
            this.labelParametersCaption.Size = new System.Drawing.Size(91, 20);
            this.labelParametersCaption.TabIndex = 21;
            this.labelParametersCaption.Text = "Parameters";
            // 
            // dataGridViewEventsParams
            // 
            this.dataGridViewEventsParams.AllowUserToAddRows = false;
            this.dataGridViewEventsParams.AllowUserToDeleteRows = false;
            this.dataGridViewEventsParams.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewEventsParams.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewEventsParams.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewEventsParams.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewEventsParams.EnableHeadersVisualStyles = false;
            this.dataGridViewEventsParams.Location = new System.Drawing.Point(0, 235);
            this.dataGridViewEventsParams.MultiSelect = false;
            this.dataGridViewEventsParams.Name = "dataGridViewEventsParams";
            this.dataGridViewEventsParams.RowHeadersVisible = false;
            this.dataGridViewEventsParams.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewEventsParams.ShowCellToolTips = false;
            this.dataGridViewEventsParams.ShowEditingIcon = false;
            this.dataGridViewEventsParams.Size = new System.Drawing.Size(777, 158);
            this.dataGridViewEventsParams.TabIndex = 20;
            // 
            // dataGridViewEvents
            // 
            this.dataGridViewEvents.AllowUserToAddRows = false;
            this.dataGridViewEvents.AllowUserToDeleteRows = false;
            this.dataGridViewEvents.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewEvents.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewEvents.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewEvents.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewEvents.EnableHeadersVisualStyles = false;
            this.dataGridViewEvents.Location = new System.Drawing.Point(0, 32);
            this.dataGridViewEvents.MultiSelect = false;
            this.dataGridViewEvents.Name = "dataGridViewEvents";
            this.dataGridViewEvents.RowHeadersVisible = false;
            this.dataGridViewEvents.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewEvents.ShowCellToolTips = false;
            this.dataGridViewEvents.ShowEditingIcon = false;
            this.dataGridViewEvents.Size = new System.Drawing.Size(777, 167);
            this.dataGridViewEvents.TabIndex = 19;
            this.dataGridViewEvents.SelectionChanged += new System.EventHandler(this.dataGridViewMethods_SelectionChanged);
            // 
            // EventsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.textBoxMethodFilter);
            this.Controls.Add(this.labelFilterCaption);
            this.Controls.Add(this.labelParametersCaption);
            this.Controls.Add(this.dataGridViewEventsParams);
            this.Controls.Add(this.dataGridViewEvents);
            this.Name = "EventsControl";
            this.Size = new System.Drawing.Size(777, 399);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEventsParams)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewEvents)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxMethodFilter;
        private System.Windows.Forms.Label labelFilterCaption;
        private System.Windows.Forms.Label labelParametersCaption;
        private System.Windows.Forms.DataGridView dataGridViewEventsParams;
        private System.Windows.Forms.DataGridView dataGridViewEvents;
    }
}
