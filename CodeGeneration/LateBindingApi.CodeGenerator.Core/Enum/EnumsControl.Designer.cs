namespace LateBindingApi.CodeGenerator.Core
{
    partial class EnumsControl
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
            this.labelEnumsControl = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelEnumsControl
            // 
            this.labelEnumsControl.AutoSize = true;
            this.labelEnumsControl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelEnumsControl.Location = new System.Drawing.Point(30, 25);
            this.labelEnumsControl.Name = "labelEnumsControl";
            this.labelEnumsControl.Size = new System.Drawing.Size(143, 20);
            this.labelEnumsControl.TabIndex = 4;
            this.labelEnumsControl.Text = "labelEnumsControl";
            // 
            // EnumsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelEnumsControl);
            this.Name = "EnumsControl";
            this.Size = new System.Drawing.Size(705, 393);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelEnumsControl;
    }
}
