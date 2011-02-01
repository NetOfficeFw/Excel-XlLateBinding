namespace LateBindingApi.CodeGenerator.Core.Visual
{
    partial class ComponentTab
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
            this.tabControlComponent = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.methodsGridMain = new LateBindingApi.CodeGenerator.Core.Visual.MethodsGrid();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.propertiesGridMain = new LateBindingApi.CodeGenerator.Core.Visual.PropertiesGrid();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.eventsGridMain = new LateBindingApi.CodeGenerator.Core.Visual.EventsGrid();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.xmlControlMain = new LateBindingApi.CodeGenerator.Core.Visual.XmlControl();
            this.itemEntitiesGridMain = new LateBindingApi.CodeGenerator.Core.Visual.ItemEntitiesGrid();
            this.tabControlComponent.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage5.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControlComponent
            // 
            this.tabControlComponent.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControlComponent.Controls.Add(this.tabPage1);
            this.tabControlComponent.Controls.Add(this.tabPage2);
            this.tabControlComponent.Controls.Add(this.tabPage3);
            this.tabControlComponent.Controls.Add(this.tabPage5);
            this.tabControlComponent.Controls.Add(this.tabPage4);
            this.tabControlComponent.Location = new System.Drawing.Point(0, 0);
            this.tabControlComponent.Name = "tabControlComponent";
            this.tabControlComponent.SelectedIndex = 0;
            this.tabControlComponent.Size = new System.Drawing.Size(524, 447);
            this.tabControlComponent.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.itemEntitiesGridMain);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(516, 421);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Attributes";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.methodsGridMain);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(516, 421);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Methods";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // methodsGridMain
            // 
            this.methodsGridMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.methodsGridMain.Location = new System.Drawing.Point(0, 0);
            this.methodsGridMain.Name = "methodsGridMain";
            this.methodsGridMain.Size = new System.Drawing.Size(510, 415);
            this.methodsGridMain.TabIndex = 0;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.propertiesGridMain);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(516, 421);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Properties";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // propertiesGridMain
            // 
            this.propertiesGridMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.propertiesGridMain.Location = new System.Drawing.Point(0, 0);
            this.propertiesGridMain.Name = "propertiesGridMain";
            this.propertiesGridMain.Size = new System.Drawing.Size(510, 415);
            this.propertiesGridMain.TabIndex = 0;
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.eventsGridMain);
            this.tabPage5.Location = new System.Drawing.Point(4, 22);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage5.Size = new System.Drawing.Size(516, 421);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.Text = "Events";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // eventsGridMain
            // 
            this.eventsGridMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.eventsGridMain.Location = new System.Drawing.Point(0, 0);
            this.eventsGridMain.Name = "eventsGridMain";
            this.eventsGridMain.Size = new System.Drawing.Size(510, 415);
            this.eventsGridMain.TabIndex = 0;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.xmlControlMain);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(516, 421);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Xml";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // xmlControlMain
            // 
            this.xmlControlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.xmlControlMain.Location = new System.Drawing.Point(0, 0);
            this.xmlControlMain.Name = "xmlControlMain";
            this.xmlControlMain.Size = new System.Drawing.Size(516, 421);
            this.xmlControlMain.TabIndex = 0;
            // 
            // itemEntitiesGridMain
            // 
            this.itemEntitiesGridMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.itemEntitiesGridMain.Location = new System.Drawing.Point(3, 3);
            this.itemEntitiesGridMain.Name = "itemEntitiesGridMain";
            this.itemEntitiesGridMain.Size = new System.Drawing.Size(510, 415);
            this.itemEntitiesGridMain.TabIndex = 0;
            // 
            // ComponentTab
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControlComponent);
            this.Name = "ComponentTab";
            this.Size = new System.Drawing.Size(524, 447);
            this.tabControlComponent.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.tabPage5.ResumeLayout(false);
            this.tabPage4.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControlComponent;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.TabPage tabPage2;
        private MethodsGrid methodsGridMain;
        private PropertiesGrid propertiesGridMain;
        private XmlControl xmlControlMain;
        private System.Windows.Forms.TabPage tabPage5;
        private EventsGrid eventsGridMain;
        private ItemEntitiesGrid itemEntitiesGridMain;
    }
}
