namespace LateBindingApi.CodeGenerator.Core.Visual
{
    partial class ComponentTreeView
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
            this.textBoxFilter = new System.Windows.Forms.TextBox();
            this.labelFilterCaption = new System.Windows.Forms.Label();
            this.treeViewComponents = new System.Windows.Forms.TreeView();
            this.SuspendLayout();
            // 
            // textBoxFilter
            // 
            this.textBoxFilter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxFilter.BackColor = System.Drawing.Color.DarkKhaki;
            this.textBoxFilter.Location = new System.Drawing.Point(41, 0);
            this.textBoxFilter.Name = "textBoxFilter";
            this.textBoxFilter.Size = new System.Drawing.Size(400, 20);
            this.textBoxFilter.TabIndex = 6;
            this.textBoxFilter.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBoxFilter_KeyDown);
            // 
            // labelFilterCaption
            // 
            this.labelFilterCaption.AutoSize = true;
            this.labelFilterCaption.Location = new System.Drawing.Point(6, 3);
            this.labelFilterCaption.Name = "labelFilterCaption";
            this.labelFilterCaption.Size = new System.Drawing.Size(29, 13);
            this.labelFilterCaption.TabIndex = 5;
            this.labelFilterCaption.Text = "Filter";
            // 
            // treeViewComponents
            // 
            this.treeViewComponents.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.treeViewComponents.BackColor = System.Drawing.Color.DarkKhaki;
            this.treeViewComponents.HideSelection = false;
            this.treeViewComponents.Location = new System.Drawing.Point(0, 26);
            this.treeViewComponents.Name = "treeViewComponents";
            this.treeViewComponents.Size = new System.Drawing.Size(441, 499);
            this.treeViewComponents.TabIndex = 4;
            this.treeViewComponents.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeViewComponents_AfterSelect);
            // 
            // ComponentTreeView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.textBoxFilter);
            this.Controls.Add(this.labelFilterCaption);
            this.Controls.Add(this.treeViewComponents);
            this.Name = "ComponentTreeView";
            this.Size = new System.Drawing.Size(441, 525);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxFilter;
        private System.Windows.Forms.Label labelFilterCaption;
        private System.Windows.Forms.TreeView treeViewComponents;
    }
}
