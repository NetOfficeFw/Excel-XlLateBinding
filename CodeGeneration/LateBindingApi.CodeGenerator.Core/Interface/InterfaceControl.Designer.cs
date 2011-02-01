namespace LateBindingApi.CodeGenerator.Core
{
    partial class InterfaceControl
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
            this.tabControlCoClass = new System.Windows.Forms.TabControl();
            this.tabPageInterface = new System.Windows.Forms.TabPage();
            this.detailsControlMain = new LateBindingApi.CodeGenerator.Core.DetailsControl();
            this.tabPageProperties = new System.Windows.Forms.TabPage();
            this.propertiesControlMain = new LateBindingApi.CodeGenerator.Core.PropertiesControl();
            this.tabPageMethods = new System.Windows.Forms.TabPage();
            this.methodsControlMain = new LateBindingApi.CodeGenerator.Core.MethodsControl();
            this.tabControlCoClass.SuspendLayout();
            this.tabPageInterface.SuspendLayout();
            this.tabPageProperties.SuspendLayout();
            this.tabPageMethods.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControlCoClass
            // 
            this.tabControlCoClass.Controls.Add(this.tabPageInterface);
            this.tabControlCoClass.Controls.Add(this.tabPageProperties);
            this.tabControlCoClass.Controls.Add(this.tabPageMethods);
            this.tabControlCoClass.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlCoClass.Location = new System.Drawing.Point(0, 0);
            this.tabControlCoClass.Name = "tabControlCoClass";
            this.tabControlCoClass.SelectedIndex = 0;
            this.tabControlCoClass.Size = new System.Drawing.Size(705, 393);
            this.tabControlCoClass.TabIndex = 2;
            // 
            // tabPageInterface
            // 
            this.tabPageInterface.Controls.Add(this.detailsControlMain);
            this.tabPageInterface.Location = new System.Drawing.Point(4, 22);
            this.tabPageInterface.Name = "tabPageInterface";
            this.tabPageInterface.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageInterface.Size = new System.Drawing.Size(697, 367);
            this.tabPageInterface.TabIndex = 3;
            this.tabPageInterface.Text = "Interface";
            this.tabPageInterface.UseVisualStyleBackColor = true;
            // 
            // detailsControlMain
            // 
            this.detailsControlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.detailsControlMain.Location = new System.Drawing.Point(3, 3);
            this.detailsControlMain.Name = "detailsControlMain";
            this.detailsControlMain.Size = new System.Drawing.Size(691, 361);
            this.detailsControlMain.TabIndex = 0;
            // 
            // tabPageProperties
            // 
            this.tabPageProperties.Controls.Add(this.propertiesControlMain);
            this.tabPageProperties.Location = new System.Drawing.Point(4, 22);
            this.tabPageProperties.Name = "tabPageProperties";
            this.tabPageProperties.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageProperties.Size = new System.Drawing.Size(697, 367);
            this.tabPageProperties.TabIndex = 0;
            this.tabPageProperties.Text = "Properties";
            this.tabPageProperties.UseVisualStyleBackColor = true;
            // 
            // propertiesControlMain
            // 
            this.propertiesControlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.propertiesControlMain.Location = new System.Drawing.Point(3, 3);
            this.propertiesControlMain.Name = "propertiesControlMain";
            this.propertiesControlMain.Size = new System.Drawing.Size(691, 361);
            this.propertiesControlMain.TabIndex = 0;
            // 
            // tabPageMethods
            // 
            this.tabPageMethods.Controls.Add(this.methodsControlMain);
            this.tabPageMethods.Location = new System.Drawing.Point(4, 22);
            this.tabPageMethods.Name = "tabPageMethods";
            this.tabPageMethods.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageMethods.Size = new System.Drawing.Size(697, 367);
            this.tabPageMethods.TabIndex = 1;
            this.tabPageMethods.Text = "Methods";
            this.tabPageMethods.UseVisualStyleBackColor = true;
            // 
            // methodsControlMain
            // 
            this.methodsControlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.methodsControlMain.Location = new System.Drawing.Point(3, 3);
            this.methodsControlMain.Name = "methodsControlMain";
            this.methodsControlMain.Size = new System.Drawing.Size(691, 361);
            this.methodsControlMain.TabIndex = 0;
            // 
            // InterfaceControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControlCoClass);
            this.Name = "InterfaceControl";
            this.Size = new System.Drawing.Size(705, 393);
            this.tabControlCoClass.ResumeLayout(false);
            this.tabPageInterface.ResumeLayout(false);
            this.tabPageProperties.ResumeLayout(false);
            this.tabPageMethods.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControlCoClass;
        private System.Windows.Forms.TabPage tabPageInterface;
        private DetailsControl detailsControlMain;
        private System.Windows.Forms.TabPage tabPageProperties;
        private PropertiesControl propertiesControlMain;
        private System.Windows.Forms.TabPage tabPageMethods;
        private MethodsControl methodsControlMain;
    }
}
