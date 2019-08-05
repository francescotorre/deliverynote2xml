namespace deliverynote2xml
{
    partial class formDeliverynote2xml
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(formDeliverynote2xml));
            this.txtRtfSource = new System.Windows.Forms.TextBox();
            this.btnOpenRtf = new System.Windows.Forms.Button();
            this.txtXmlDestination = new System.Windows.Forms.TextBox();
            this.btnOpenXml = new System.Windows.Forms.Button();
            this.btnConvert = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtCustomerDataFilePath = new System.Windows.Forms.TextBox();
            this.btnChangeCustomerDataFilePath = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtRtfSource
            // 
            this.txtRtfSource.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtRtfSource.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRtfSource.Location = new System.Drawing.Point(216, 30);
            this.txtRtfSource.Multiline = true;
            this.txtRtfSource.Name = "txtRtfSource";
            this.txtRtfSource.Size = new System.Drawing.Size(367, 31);
            this.txtRtfSource.TabIndex = 1;
            this.txtRtfSource.TabStop = false;
            this.txtRtfSource.WordWrap = false;
            this.txtRtfSource.ReadOnlyChanged += new System.EventHandler(this.txtRtfSource_ReadOnlyChanged);
            // 
            // btnOpenRtf
            // 
            this.btnOpenRtf.Location = new System.Drawing.Point(46, 30);
            this.btnOpenRtf.Name = "btnOpenRtf";
            this.btnOpenRtf.Size = new System.Drawing.Size(153, 31);
            this.btnOpenRtf.TabIndex = 0;
            this.btnOpenRtf.Text = "Open RTF file";
            this.btnOpenRtf.UseVisualStyleBackColor = true;
            this.btnOpenRtf.Click += new System.EventHandler(this.btnOpenRtf_Click);
            // 
            // txtXmlDestination
            // 
            this.txtXmlDestination.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtXmlDestination.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtXmlDestination.Location = new System.Drawing.Point(216, 83);
            this.txtXmlDestination.Multiline = true;
            this.txtXmlDestination.Name = "txtXmlDestination";
            this.txtXmlDestination.Size = new System.Drawing.Size(367, 31);
            this.txtXmlDestination.TabIndex = 3;
            this.txtXmlDestination.TabStop = false;
            this.txtXmlDestination.WordWrap = false;
            this.txtXmlDestination.ReadOnlyChanged += new System.EventHandler(this.txtXmlDestination_ReadOnlyChanged);
            // 
            // btnOpenXml
            // 
            this.btnOpenXml.Enabled = false;
            this.btnOpenXml.Location = new System.Drawing.Point(46, 83);
            this.btnOpenXml.Name = "btnOpenXml";
            this.btnOpenXml.Size = new System.Drawing.Size(153, 31);
            this.btnOpenXml.TabIndex = 2;
            this.btnOpenXml.Text = "Save XML file as...";
            this.btnOpenXml.UseVisualStyleBackColor = true;
            this.btnOpenXml.Click += new System.EventHandler(this.btnOpenXml_Click);
            // 
            // btnConvert
            // 
            this.btnConvert.Enabled = false;
            this.btnConvert.Location = new System.Drawing.Point(46, 234);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(537, 42);
            this.btnConvert.TabIndex = 4;
            this.btnConvert.Text = "CONVERT";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.txtCustomerDataFilePath);
            this.panel1.Controls.Add(this.btnChangeCustomerDataFilePath);
            this.panel1.Location = new System.Drawing.Point(46, 132);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(537, 76);
            this.panel1.TabIndex = 3;
            // 
            // txtCustomerRegistry
            // 
            this.txtCustomerDataFilePath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtCustomerDataFilePath.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCustomerDataFilePath.Location = new System.Drawing.Point(23, 22);
            this.txtCustomerDataFilePath.Multiline = true;
            this.txtCustomerDataFilePath.Name = "txtCustomerRegistry";
            this.txtCustomerDataFilePath.Size = new System.Drawing.Size(367, 31);
            this.txtCustomerDataFilePath.TabIndex = 0;
            this.txtCustomerDataFilePath.TabStop = false;
            this.txtCustomerDataFilePath.WordWrap = false;
            this.txtCustomerDataFilePath.ReadOnlyChanged += new System.EventHandler(this.txtCustomerRegistry_ReadOnlyChanged);
            // 
            // btnChangeRegistryPath
            // 
            this.btnChangeCustomerDataFilePath.Location = new System.Drawing.Point(406, 22);
            this.btnChangeCustomerDataFilePath.Name = "btnChangeRegistryPath";
            this.btnChangeCustomerDataFilePath.Size = new System.Drawing.Size(104, 31);
            this.btnChangeCustomerDataFilePath.TabIndex = 0;
            this.btnChangeCustomerDataFilePath.Text = "Change";
            this.btnChangeCustomerDataFilePath.UseVisualStyleBackColor = true;
            this.btnChangeCustomerDataFilePath.Click += new System.EventHandler(this.btnChangeCustomerDataFilePath_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(67, 123);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(169, 17);
            this.label1.TabIndex = 6;
            this.label1.Text = "Path to customer data file";
            // 
            // formDeliverynote2xml
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(629, 297);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnConvert);
            this.Controls.Add(this.btnOpenXml);
            this.Controls.Add(this.btnOpenRtf);
            this.Controls.Add(this.txtXmlDestination);
            this.Controls.Add(this.txtRtfSource);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "formDeliverynote2xml";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "DeliveryNote2XML | RTF to XML file converter";
            this.Load += new System.EventHandler(this.formDeliverynote2xml_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtRtfSource;
        private System.Windows.Forms.Button btnOpenRtf;
        private System.Windows.Forms.TextBox txtXmlDestination;
        private System.Windows.Forms.Button btnOpenXml;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtCustomerDataFilePath;
        private System.Windows.Forms.Button btnChangeCustomerDataFilePath;
    }
}