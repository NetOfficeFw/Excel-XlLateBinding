namespace LateBindingApi.CodeGenerator
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
            this.tabControlInterface = new System.Windows.Forms.TabControl();
            this.tabPageCoClass = new System.Windows.Forms.TabPage();
            this.dataGridViewInterfaceInterfaces = new System.Windows.Forms.DataGridView();
            this.labelInterfacesCaption = new System.Windows.Forms.Label();
            this.labelDefinedCaption = new System.Windows.Forms.Label();
            this.dataGridViewInterfaceComponents = new System.Windows.Forms.DataGridView();
            this.tabPageMethods = new System.Windows.Forms.TabPage();
            this.methodsControlMain = new LateBindingApi.CodeGenerator.MethodsControl();
            this.textBoxGUID = new System.Windows.Forms.TextBox();
            this.labeGuidCaption = new System.Windows.Forms.Label();
            this.tabControlInterface.SuspendLayout();
            this.tabPageCoClass.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewInterfaceInterfaces)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewInterfaceComponents)).BeginInit();
            this.tabPageMethods.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControlInterface
            // 
            this.tabControlInterface.Controls.Add(this.tabPageCoClass);
            this.tabControlInterface.Controls.Add(this.tabPageMethods);
            this.tabControlInterface.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlInterface.Location = new System.Drawing.Point(0, 0);
            this.tabControlInterface.Name = "tabControlInterface";
            this.tabControlInterface.SelectedIndex = 0;
            this.tabControlInterface.Size = new System.Drawing.Size(686, 342);
            this.tabControlInterface.TabIndex = 1;
            // 
            // tabPageCoClass
            // 
            this.tabPageCoClass.Controls.Add(this.textBoxGUID);
            this.tabPageCoClass.Controls.Add(this.labeGuidCaption);
            this.tabPageCoClass.Controls.Add(this.dataGridViewInterfaceInterfaces);
            this.tabPageCoClass.Controls.Add(this.labelInterfacesCaption);
            this.tabPageCoClass.Controls.Add(this.labelDefinedCaption);
            this.tabPageCoClass.Controls.Add(this.dataGridViewInterfaceComponents);
            this.tabPageCoClass.Location = new System.Drawing.Point(4, 22);
            this.tabPageCoClass.Name = "tabPageCoClass";
            this.tabPageCoClass.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageCoClass.Size = new System.Drawing.Size(678, 316);
            this.tabPageCoClass.TabIndex = 3;
            this.tabPageCoClass.Text = "Interface";
            this.tabPageCoClass.UseVisualStyleBackColor = true;
            // 
            // dataGridViewInterfaceInterfaces
            // 
            this.dataGridViewInterfaceInterfaces.AllowUserToAddRows = false;
            this.dataGridViewInterfaceInterfaces.AllowUserToDeleteRows = false;
            this.dataGridViewInterfaceInterfaces.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewInterfaceInterfaces.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewInterfaceInterfaces.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewInterfaceInterfaces.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewInterfaceInterfaces.EnableHeadersVisualStyles = false;
            this.dataGridViewInterfaceInterfaces.Location = new System.Drawing.Point(1, 168);
            this.dataGridViewInterfaceInterfaces.MultiSelect = false;
            this.dataGridViewInterfaceInterfaces.Name = "dataGridViewInterfaceInterfaces";
            this.dataGridViewInterfaceInterfaces.RowHeadersVisible = false;
            this.dataGridViewInterfaceInterfaces.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewInterfaceInterfaces.ShowCellToolTips = false;
            this.dataGridViewInterfaceInterfaces.ShowEditingIcon = false;
            this.dataGridViewInterfaceInterfaces.Size = new System.Drawing.Size(678, 149);
            this.dataGridViewInterfaceInterfaces.TabIndex = 9;
            // 
            // labelInterfacesCaption
            // 
            this.labelInterfacesCaption.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelInterfacesCaption.AutoSize = true;
            this.labelInterfacesCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelInterfacesCaption.Location = new System.Drawing.Point(3, 145);
            this.labelInterfacesCaption.Name = "labelInterfacesCaption";
            this.labelInterfacesCaption.Size = new System.Drawing.Size(326, 20);
            this.labelInterfacesCaption.TabIndex = 8;
            this.labelInterfacesCaption.Text = "Interface implements the following Interfaces";
            // 
            // labelDefinedCaption
            // 
            this.labelDefinedCaption.AutoSize = true;
            this.labelDefinedCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelDefinedCaption.Location = new System.Drawing.Point(3, 19);
            this.labelDefinedCaption.Name = "labelDefinedCaption";
            this.labelDefinedCaption.Size = new System.Drawing.Size(348, 20);
            this.labelDefinedCaption.TabIndex = 7;
            this.labelDefinedCaption.Text = "Interface is defined in the following Components";
            // 
            // dataGridViewInterfaceComponents
            // 
            this.dataGridViewInterfaceComponents.AllowUserToAddRows = false;
            this.dataGridViewInterfaceComponents.AllowUserToDeleteRows = false;
            this.dataGridViewInterfaceComponents.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewInterfaceComponents.BackgroundColor = System.Drawing.Color.DarkKhaki;
            this.dataGridViewInterfaceComponents.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dataGridViewInterfaceComponents.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewInterfaceComponents.EnableHeadersVisualStyles = false;
            this.dataGridViewInterfaceComponents.Location = new System.Drawing.Point(0, 42);
            this.dataGridViewInterfaceComponents.MultiSelect = false;
            this.dataGridViewInterfaceComponents.Name = "dataGridViewInterfaceComponents";
            this.dataGridViewInterfaceComponents.RowHeadersVisible = false;
            this.dataGridViewInterfaceComponents.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridViewInterfaceComponents.ShowCellToolTips = false;
            this.dataGridViewInterfaceComponents.ShowEditingIcon = false;
            this.dataGridViewInterfaceComponents.Size = new System.Drawing.Size(678, 74);
            this.dataGridViewInterfaceComponents.TabIndex = 6;
            // 
            // tabPageMethods
            // 
            this.tabPageMethods.Controls.Add(this.methodsControlMain);
            this.tabPageMethods.Location = new System.Drawing.Point(4, 22);
            this.tabPageMethods.Name = "tabPageMethods";
            this.tabPageMethods.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageMethods.Size = new System.Drawing.Size(678, 316);
            this.tabPageMethods.TabIndex = 1;
            this.tabPageMethods.Text = "Methods";
            this.tabPageMethods.UseVisualStyleBackColor = true;
            // 
            // methodsControlMain
            // 
            this.methodsControlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.methodsControlMain.Location = new System.Drawing.Point(3, 3);
            this.methodsControlMain.Name = "methodsControlMain";
            this.methodsControlMain.Size = new System.Drawing.Size(672, 310);
            this.methodsControlMain.TabIndex = 0;
            // 
            // textBoxGUID
            // 
            this.textBoxGUID.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxGUID.BackColor = System.Drawing.Color.DarkKhaki;
            this.textBoxGUID.Location = new System.Drawing.Point(420, 16);
            this.textBoxGUID.Name = "textBoxGUID";
            this.textBoxGUID.ReadOnly = true;
            this.textBoxGUID.Size = new System.Drawing.Size(235, 20);
            this.textBoxGUID.TabIndex = 11;
            // 
            // labeGuidCaption
            // 
            this.labeGuidCaption.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labeGuidCaption.AutoSize = true;
            this.labeGuidCaption.Location = new System.Drawing.Point(380, 19);
            this.labeGuidCaption.Name = "labeGuidCaption";
            this.labeGuidCaption.Size = new System.Drawing.Size(34, 13);
            this.labeGuidCaption.TabIndex = 10;
            this.labeGuidCaption.Text = "GUID";
            // 
            // InterfaceControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabControlInterface);
            this.Name = "InterfaceControl";
            this.Size = new System.Drawing.Size(686, 342);
            this.tabControlInterface.ResumeLayout(false);
            this.tabPageCoClass.ResumeLayout(false);
            this.tabPageCoClass.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewInterfaceInterfaces)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewInterfaceComponents)).EndInit();
            this.tabPageMethods.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControlInterface;
        private System.Windows.Forms.TabPage tabPageCoClass;
        private System.Windows.Forms.DataGridView dataGridViewInterfaceInterfaces;
        private System.Windows.Forms.Label labelInterfacesCaption;
        private System.Windows.Forms.Label labelDefinedCaption;
        private System.Windows.Forms.DataGridView dataGridViewInterfaceComponents;
        private System.Windows.Forms.TabPage tabPageMethods;
        private MethodsControl methodsControlMain;
        private System.Windows.Forms.TextBox textBoxGUID;
        private System.Windows.Forms.Label labeGuidCaption;
    }
}
