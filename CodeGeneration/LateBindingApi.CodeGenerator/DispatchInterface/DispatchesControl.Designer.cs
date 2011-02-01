namespace LateBindingApi.CodeGenerator
{
    partial class DispatchesControl
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
            this.labelInterfacesInfo = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelInterfacesInfo
            // 
            this.labelInterfacesInfo.AutoSize = true;
            this.labelInterfacesInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelInterfacesInfo.Location = new System.Drawing.Point(30, 25);
            this.labelInterfacesInfo.Name = "labelInterfacesInfo";
            this.labelInterfacesInfo.Size = new System.Drawing.Size(142, 20);
            this.labelInterfacesInfo.TabIndex = 4;
            this.labelInterfacesInfo.Text = "labelInterfacesInfo";
            // 
            // DispatchesControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelInterfacesInfo);
            this.Name = "DispatchesControl";
            this.Size = new System.Drawing.Size(686, 342);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelInterfacesInfo;
    }
}
