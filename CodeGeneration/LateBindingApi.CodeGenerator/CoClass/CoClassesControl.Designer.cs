﻿namespace LateBindingApi.CodeGenerator
{
    partial class CoClassesControl
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
            this.labelClassesInfo = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelClassesInfo
            // 
            this.labelClassesInfo.AutoSize = true;
            this.labelClassesInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelClassesInfo.Location = new System.Drawing.Point(30, 25);
            this.labelClassesInfo.Name = "labelClassesInfo";
            this.labelClassesInfo.Size = new System.Drawing.Size(126, 20);
            this.labelClassesInfo.TabIndex = 2;
            this.labelClassesInfo.Text = "labelClassesInfo";
            // 
            // CoClassesControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelClassesInfo);
            this.Name = "CoClassesControl";
            this.Size = new System.Drawing.Size(686, 342);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelClassesInfo;
    }
}
