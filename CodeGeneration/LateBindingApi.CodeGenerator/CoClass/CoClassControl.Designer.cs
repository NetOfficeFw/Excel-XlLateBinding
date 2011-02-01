namespace LateBindingApi.CodeGenerator
{
    partial class CoClassControl
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
            this.tabPageCoClass = new System.Windows.Forms.TabPage();
            this.dataGridViewCoClassInterfaces = new System.Windows.Forms.DataGridView();
            this.labelInterfacesCaption = new System.Windows.Forms.Label();
            this.labelDefinedCaption = new System.Windows.Forms.Label();
            this.dataGridViewCoClassComponents = new System.Windows.Forms.DataGridView();
            this.tabPageProperties = new System.Windows.Forms.TabPage();
            this.tabPageMethods = new System.Windows.Forms.TabPage();
            this.tabPageEvents = new System.Windows.Forms.TabPage();
            this.propertiesControlMain = new LateBindingApi.CodeGenerator.PropertiesControl();
            this.methodsControlMain = new LateBindingApi.CodeGenerator.MethodsControl();
            this.eventsControlMain = new LateBindingApi.CodeGenerator.EventsControl();
            this.tabControlCoClass.SuspendLayout();
            this.tabPageCoClass.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewCoClassInterfaces)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewCoClassComponents)).BeginInit();
            this.tabPageProperties.SuspendLayout();
            this.tabPageMethods.SuspendLayout();
            this.tabPageEvents.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControlCoClass
            // 
            this.tabControlCoClass.Controls.Add(this.tabPageCoClass);
            this.tabControlCoClass.Controls.Add(this.tabPageProperties);
            this.tabControlCoClass.Controls.Add(this.tabPageMethods);
            this.tabControlCoClass.Controls.Add(this.tabPageEvents);
            this.tabControlCoClass.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlCoClass.Location = new System.Drawing.Point(0, 0);
            this.tabControlCoClass.Name = "tabControlCoClass";
            this.tabControlCoClass.SelectedIndex = 0;
            this.tabControlCoClass.Size = new System.Drawing.Size(792, 399);
            this.tabControlCoClass.TabIndex = 0;
            // 
            // tabPageCoClass
            // 
            this.tabPageCoClass.Controls.Add(this.dataGridViewCoClassInterfaces);
            this.tabPageCoClass.Controls.Add(this.labelInterfacesCaption);
            this.tabPageCoClass.Controls.Add(this.labelDefinedCaption);
            this.tabPageCoClass.Controls.Add(this.dataGridViewCoClassComponents);
            this.tabPageCoClass.Location = new System.Drawing.Point(4, 22);
            this.tabPageCoClass.Name = "tabPageCoClass";
            this.tabPageCoClass.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageCoClass.Size = new System.Drawing.Size(784, 373);
            this.tabPageCoClass.TabIndex = 3;
            this.tabPageCoClass.Text = "CoClass";
            this.tabPageCoClass.UseVisualStyleBackColor = true;
            // 
            // dataGridViewCoClassInterfaces
            // 
            this.dataGridViewCoClassInterfaces.AllowUserToAddRows = false;
            this.dataGridViewCoClassInterfaces.AllowUserToDeleteRows = false;
            this.dataGridViewCoClassInterfaces.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewCoClassInterfaces.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewCoClassInterfaces.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewCoClassInterfaces.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewCoClassInterfaces.EnableHeadersVisualStyles = false;
            this.dataGridViewCoClassInterfaces.Location = new System.Drawing.Point(1, 225);
            this.dataGridViewCoClassInterfaces.MultiSelect = false;
            this.dataGridViewCoClassInterfaces.Name = "dataGridViewCoClassInterfaces";
            this.dataGridViewCoClassInterfaces.RowHeadersVisible = false;
            this.dataGridViewCoClassInterfaces.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewCoClassInterfaces.ShowCellToolTips = false;
            this.dataGridViewCoClassInterfaces.ShowEditingIcon = false;
            this.dataGridViewCoClassInterfaces.Size = new System.Drawing.Size(784, 131);
            this.dataGridViewCoClassInterfaces.TabIndex = 9;
            // 
            // labelInterfacesCaption
            // 
            this.labelInterfacesCaption.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelInterfacesCaption.AutoSize = true;
            this.labelInterfacesCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelInterfacesCaption.Location = new System.Drawing.Point(3, 202);
            this.labelInterfacesCaption.Name = "labelInterfacesCaption";
            this.labelInterfacesCaption.Size = new System.Drawing.Size(321, 20);
            this.labelInterfacesCaption.TabIndex = 8;
            this.labelInterfacesCaption.Text = "CoClass implements the following Interfaces";
            // 
            // labelDefinedCaption
            // 
            this.labelDefinedCaption.AutoSize = true;
            this.labelDefinedCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelDefinedCaption.Location = new System.Drawing.Point(3, 19);
            this.labelDefinedCaption.Name = "labelDefinedCaption";
            this.labelDefinedCaption.Size = new System.Drawing.Size(343, 20);
            this.labelDefinedCaption.TabIndex = 7;
            this.labelDefinedCaption.Text = "CoClass is defined in the following Components";
            // 
            // dataGridViewCoClassComponents
            // 
            this.dataGridViewCoClassComponents.AllowUserToAddRows = false;
            this.dataGridViewCoClassComponents.AllowUserToDeleteRows = false;
            this.dataGridViewCoClassComponents.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewCoClassComponents.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewCoClassComponents.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewCoClassComponents.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewCoClassComponents.EnableHeadersVisualStyles = false;
            this.dataGridViewCoClassComponents.Location = new System.Drawing.Point(0, 42);
            this.dataGridViewCoClassComponents.MultiSelect = false;
            this.dataGridViewCoClassComponents.Name = "dataGridViewCoClassComponents";
            this.dataGridViewCoClassComponents.RowHeadersVisible = false;
            this.dataGridViewCoClassComponents.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewCoClassComponents.ShowCellToolTips = false;
            this.dataGridViewCoClassComponents.ShowEditingIcon = false;
            this.dataGridViewCoClassComponents.Size = new System.Drawing.Size(784, 131);
            this.dataGridViewCoClassComponents.TabIndex = 6;
            // 
            // tabPageProperties
            // 
            this.tabPageProperties.Controls.Add(this.propertiesControlMain);
            this.tabPageProperties.Location = new System.Drawing.Point(4, 22);
            this.tabPageProperties.Name = "tabPageProperties";
            this.tabPageProperties.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageProperties.Size = new System.Drawing.Size(784, 373);
            this.tabPageProperties.TabIndex = 0;
            this.tabPageProperties.Text = "Properties";
            this.tabPageProperties.UseVisualStyleBackColor = true;
            // 
            // tabPageMethods
            // 
            this.tabPageMethods.Controls.Add(this.methodsControlMain);
            this.tabPageMethods.Location = new System.Drawing.Point(4, 22);
            this.tabPageMethods.Name = "tabPageMethods";
            this.tabPageMethods.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageMethods.Size = new System.Drawing.Size(784, 373);
            this.tabPageMethods.TabIndex = 1;
            this.tabPageMethods.Text = "Methods";
            this.tabPageMethods.UseVisualStyleBackColor = true;
            // 
            // tabPageEvents
            // 
            this.tabPageEvents.Controls.Add(this.eventsControlMain);
            this.tabPageEvents.Location = new System.Drawing.Point(4, 22);
            this.tabPageEvents.Name = "tabPageEvents";
            this.tabPageEvents.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageEvents.Size = new System.Drawing.Size(784, 373);
            this.tabPageEvents.TabIndex = 2;
            this.tabPageEvents.Text = "Events";
            this.tabPageEvents.UseVisualStyleBackColor = true;
            // 
            // propertiesControlMain
            // 
            this.propertiesControlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.propertiesControlMain.Location = new System.Drawing.Point(3, 3);
            this.propertiesControlMain.Name = "propertiesControlMain";
            this.propertiesControlMain.Size = new System.Drawing.Size(778, 367);
            this.propertiesControlMain.TabIndex = 0;
            // 
            // methodsControlMain
            // 
            this.methodsControlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.methodsControlMain.Location = new System.Drawing.Point(3, 3);
            this.methodsControlMain.Name = "methodsControlMain";
            this.methodsControlMain.Size = new System.Drawing.Size(778, 367);
            this.methodsControlMain.TabIndex = 0;
            // 
            // eventsControlMain
            // 
            this.eventsControlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.eventsControlMain.Location = new System.Drawing.Point(3, 3);
            this.eventsControlMain.Name = "eventsControlMain";
            this.eventsControlMain.Size = new System.Drawing.Size(778, 367);
            this.eventsControlMain.TabIndex = 0;
            // 
            // CoClassControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControlCoClass);
            this.Name = "CoClassControl";
            this.Size = new System.Drawing.Size(792, 399);
            this.tabControlCoClass.ResumeLayout(false);
            this.tabPageCoClass.ResumeLayout(false);
            this.tabPageCoClass.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewCoClassInterfaces)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewCoClassComponents)).EndInit();
            this.tabPageProperties.ResumeLayout(false);
            this.tabPageMethods.ResumeLayout(false);
            this.tabPageEvents.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControlCoClass;
        private System.Windows.Forms.TabPage tabPageProperties;
        private System.Windows.Forms.TabPage tabPageMethods;
        private System.Windows.Forms.TabPage tabPageEvents;
        private MethodsControl methodsControlMain;
        private System.Windows.Forms.TabPage tabPageCoClass;
        private System.Windows.Forms.DataGridView dataGridViewCoClassComponents;
        private System.Windows.Forms.Label labelDefinedCaption;
        private System.Windows.Forms.Label labelInterfacesCaption;
        private System.Windows.Forms.DataGridView dataGridViewCoClassInterfaces;
        private PropertiesControl propertiesControlMain;
        private EventsControl eventsControlMain;
    }
}
