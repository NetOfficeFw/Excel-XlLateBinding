namespace LateBindingApi.CodeGenerator
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
            this.labelEnumsInfo = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelEnumsInfo
            // 
            this.labelEnumsInfo.AutoSize = true;
            this.labelEnumsInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelEnumsInfo.Location = new System.Drawing.Point(30, 25);
            this.labelEnumsInfo.Name = "labelEnumsInfo";
            this.labelEnumsInfo.Size = new System.Drawing.Size(120, 20);
            this.labelEnumsInfo.TabIndex = 1;
            this.labelEnumsInfo.Text = "labelEnumsInfo";
            // 
            // EnumsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelEnumsInfo);
            this.Name = "EnumsControl";
            this.Size = new System.Drawing.Size(720, 383);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelEnumsInfo;


    }
}
