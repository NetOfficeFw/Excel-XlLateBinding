namespace LateBindingApi.CodeGenerator.Core.Visual
{
    partial class ItemEntitiesGrid
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.refInterfacesGrid = new LateBindingApi.CodeGenerator.Core.Visual.RefInterfacesGrid();
            this.refComponentsGrid = new LateBindingApi.CodeGenerator.Core.Visual.RefComponentsGrid();
            this.attributesGrid = new LateBindingApi.CodeGenerator.Core.Visual.AttributesGrid();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(17, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 16);
            this.label1.TabIndex = 2;
            this.label1.Text = "Attributes";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(17, 153);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(147, 16);
            this.label2.TabIndex = 3;
            this.label2.Text = "Defined in Components";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(17, 295);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(120, 16);
            this.label3.TabIndex = 4;
            this.label3.Text = "Inherited Interfaces";
            // 
            // refInterfacesGrid
            // 
            this.refInterfacesGrid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.refInterfacesGrid.Location = new System.Drawing.Point(0, 314);
            this.refInterfacesGrid.Name = "refInterfacesGrid";
            this.refInterfacesGrid.Size = new System.Drawing.Size(710, 118);
            this.refInterfacesGrid.TabIndex = 5;
            // 
            // refComponentsGrid
            // 
            this.refComponentsGrid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.refComponentsGrid.Location = new System.Drawing.Point(0, 172);
            this.refComponentsGrid.Name = "refComponentsGrid";
            this.refComponentsGrid.Size = new System.Drawing.Size(710, 118);
            this.refComponentsGrid.TabIndex = 1;
            // 
            // attributesGrid
            // 
            this.attributesGrid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.attributesGrid.Location = new System.Drawing.Point(0, 23);
            this.attributesGrid.Name = "attributesGrid";
            this.attributesGrid.Size = new System.Drawing.Size(710, 118);
            this.attributesGrid.TabIndex = 0;
            // 
            // ItemEntitiesGrid
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.refInterfacesGrid);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.refComponentsGrid);
            this.Controls.Add(this.attributesGrid);
            this.Name = "ItemEntitiesGrid";
            this.Size = new System.Drawing.Size(713, 435);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private AttributesGrid attributesGrid;
        private RefComponentsGrid refComponentsGrid;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private RefInterfacesGrid refInterfacesGrid;
    }
}
