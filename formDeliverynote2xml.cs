using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using deliverynote2xml.tags;
using System.IO;


namespace deliverynote2xml
{
    public partial class formDeliverynote2xml : Form
    {
        public formDeliverynote2xml()
        {
            InitializeComponent();
        }

        private void formDeliverynote2xml_Load(object sender, EventArgs e)
        {
            txtRtfSource.ReadOnly = true;
            txtXmlDestination.ReadOnly = true;
            txtCustomerDataFilePath.ReadOnly = true;

            CodeManager.CustomerDataFilePath = string.Format(@"{0}\{1}",
                Application.StartupPath, @"customerdata\customers.xlsx");

            txtCustomerDataFilePath.Text = string.Format(@"{0}", CodeManager.CustomerDataFilePath);

        }

        private void txtRtfSource_ReadOnlyChanged(object sender, EventArgs e)
        {
            if (txtRtfSource.ReadOnly)
            {
                txtRtfSource.BackColor = Color.FromKnownColor(KnownColor.White);
            }
        }

        private void txtXmlDestination_ReadOnlyChanged(object sender, EventArgs e)
        {
            if (txtXmlDestination.ReadOnly)
            {
                txtXmlDestination.BackColor = Color.FromKnownColor(KnownColor.White);
            }
        }


        private void txtCustomerRegistry_ReadOnlyChanged(object sender, EventArgs e)
        {
            if (txtCustomerDataFilePath.ReadOnly)
            {
                txtCustomerDataFilePath.BackColor = Color.FromKnownColor(KnownColor.White);
            }
        }

        private void btnOpenRtf_Click(object sender, EventArgs e)
        {

            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "Rich Text Format|*.rtf";

            opf.InitialDirectory =
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop);


            if (opf.ShowDialog() == DialogResult.OK)
            {
                string RTFpath = opf.FileName;

                if (!string.IsNullOrWhiteSpace(RTFpath))
                {
                    if (!RTFpath.ToLower().EndsWith(".rtf"))
                    {
                        MessageBox.Show("Please, select an RTF file as the source file.\nOnly files with .rtf extension are allowed",
                            "Invalid file type",
                                MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);


                    }
                    else
                    {
                        txtRtfSource.Text = RTFpath;

                        txtXmlDestination.Text
                        = string.Format(@"{0}\{1}.xml",
                        Path.GetDirectoryName(RTFpath),
                            Path.GetFileNameWithoutExtension(RTFpath));

                        btnOpenXml.Enabled = true;
                        btnConvert.Enabled = true;
                    }

                } // End if (!string.IsNullOrWhiteSpace(RTFpath))


            } // End if (opf.ShowDialog() == DialogResult.OK)

        }


        private void btnOpenXml_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "XML Documents|*.xml";

            sfd.InitialDirectory =
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop);


            if (sfd.ShowDialog() == DialogResult.OK)
            {
                string XMLpath = sfd.FileName;

                if (!string.IsNullOrWhiteSpace(XMLpath))
                {
                    if (!XMLpath.ToLower().EndsWith(".xml"))
                    {
                        MessageBox.Show("Please, select an XML file as the destination file.\nOnly files with .xml extension are allowed.", 
                            "Invalid file type", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        txtXmlDestination.Text = XMLpath;
                    }

                } // End if (!string.IsNullOrWhiteSpace(XMLpath))

            } // End if (sfd.ShowDialog() == DialogResult.OK)
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {

            if (!File.Exists(txtCustomerDataFilePath.Text))
            {
                MessageBox.Show(string.Format("Cannot find the customer data file {0}.\n\nPlease, make sure the path is correct and try again.",
                    txtCustomerDataFilePath.Text), "Customer data file not found",
                   MessageBoxButtons.OK,
                       MessageBoxIcon.Warning);
            }
            else if (CodeManager.rtfDocumentIsAlreadyOpen(txtRtfSource.Text))
            {
                MessageBox.Show(string.Format("Impossible to complete the task because the selected file: {0}, is currently opened in Microsoft Word.\n\nPlease, close Microsoft Word or the file and try again.",
                    txtRtfSource.Text), "RTF document is currently open",
                   MessageBoxButtons.OK,
                       MessageBoxIcon.Warning);
            }
            else
            {

                try
                {
                    toggleGuiState(false);

                    RtfParser parser = new RtfParser(@txtRtfSource.Text);

                    DeliveryNotes deliveryNotesXml = new DeliveryNotes();

                    //<?xml version="1.0"?>
                    //<DeliveryNotes 
                    //  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                    //  xmlns:xsd="http://www.w3.org/2001/XMLSchema"> 

                    //<DeliveryNotes>
                    //  <Company>
                    deliveryNotesXml.Company.Name = parser.Company["Name"].ToString();

                    //    <Address><Address/>
                    deliveryNotesXml.Company.Address = parser.Company["Address"].ToString();

                    //    <Postcode></Postcode>
                    deliveryNotesXml.Company.Zip = parser.Company["Zip"].ToString();

                    //    <City></City>
                    deliveryNotesXml.Company.City = parser.Company["City"].ToString();

                    //    <Province></Province>
                    deliveryNotesXml.Company.State = parser.Company["State"].ToString();

                    //    <Country></Country>
                    deliveryNotesXml.Company.Country = parser.Company["Country"].ToString();

                    //    <FiscalCode><FiscalCode/>
                    //fattDocumentXml.Company.FiscalCode = parser.Company["FiscalCode"].ToString();

                    //    <VatCode></VatCode>
                    //  </Company>
                    //fattDocumentXml.Company.VatCode = parser.Company["VatCode"].ToString();


                    //  <Documents>
                    foreach (Dictionary<string, string> d in parser.Documents)
                    {
                        // <Document>
                        Document document = new Document();

                        //  <CustomerCode></CustomerCode>
                        document.CustomerCode = d["CustomerCode"];

                        //  <CustomerName><CustomerName/>
                        document.SoldToName = d["CustomerName"];

                        //  <DeliveryName></DeliveryName>
                        document.ShipToName = d["DeliveryName"];

                        //  <DeliveryAddress></DeliveryAddress>
                        document.ShipToAddress = d["DeliveryAddress"];

                        //  <DeliveryPostcode></DeliveryPostcode>
                        document.ShipToZip = d["DeliveryZip"];

                        //  <DeliveryCity></DeliveryCity>
                        document.ShipToCity = d["DeliveryCity"];

                        //  <DeliveryProvince></DeliveryProvince>
                        document.ShipToState = d["DeliveryState"];

                        //  <Date></Date>
                        document.Date = d["Date"];

                        //  <Number></Number>
                        document.Number = d["Number"];
                        //document.TransportReason = d["TransportReason"];
                        //document.TransportDateTime = d["TransportDateTime"];

                        bool messageAlreadyShown = false;

                        //  <Rows>
                        foreach (Dictionary<string, string> r in parser.Rows)
                        {

                            //<Row>
                            Row row = new Row();

                            //  <Code></Code>
                            row.Code = r["Code"];

                            //  <Description></Description>
                            row.Description = r["Description"];

                            //  <Qty></Qty>
                            row.Qty = r["Qty"].Split(',')[0];


                            if (r["Discounts"] != string.Empty)
                            {
                                if (r["Discounts"] == "CUSTOMER_NOT_FOUND")
                                {
                                    if (!messageAlreadyShown)
                                    {
                                        MessageBox.Show("Warning: could not read discount for this customer because he/she was not found in the customer data file. \n\nThe XML file will be generated anyway.", 
                                           "Customer not found",
                                               MessageBoxButtons.OK,
                                                   MessageBoxIcon.Warning);

                                        messageAlreadyShown = true;
                                    }
                                }
                                else
                                {
                                    //<"Discounts></"Discounts>
                                    row.Discounts = r["Discounts"];
                                }
                            }

                            //  </Row>
                            //</Rows>
                            document.Rows.Add(row);
                        }

                        //</Document>
                        deliveryNotesXml.Documents.Add(document);
                    }

                    if (CodeManager.generateDeliveryNoteXml(deliveryNotesXml, txtXmlDestination.Text))
                    {
                        MessageBox.Show("XML file succesfully created",
                            "Conversion finished",
                                MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);

                        toggleGuiState(true);
                    }
                    else
                    {
                        MessageBox.Show("An error has occurred while saving the XML file. The application will now close.",
                        "Error",
                            MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                        this.Close();
                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show(
                        string.Format("The following error has occurred: {0}. The application will now close.",
                            ex.Message),
                                "Error",
                                    MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);

                    this.Close();

                }

            } // End if (CodeManager.rtfDocumentIsAlreadyOpen(txtRtfSource.Text))







        } // End private void btnConvert_Click(object sender, EventArgs e) 

        private void btnChangeCustomerDataFilePath_Click(object sender, EventArgs e)
        {
            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter
                = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls";

            opf.InitialDirectory =
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop);


            if (opf.ShowDialog() == DialogResult.OK)
            {
                string workbookPath = opf.FileName;

                if (!string.IsNullOrWhiteSpace(workbookPath))
                {
                    if (!workbookPath.EndsWith(".xlsx") && !workbookPath.EndsWith(".xls"))
                    {
                        MessageBox.Show("Please, select only XLSX and XLS files.\n\nAt this time, only files with .xsl and .xslx extensions are supported when importing a customer data file.",
                            "Invalid file type",
                                MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);


                    }
                    else
                    {
                        CodeManager.CustomerDataFilePath = string.Format(@"{0}", workbookPath);
                        txtCustomerDataFilePath.Text = workbookPath;
                    }

                } // End if (!string.IsNullOrWhiteSpace(RTFpath))


            } // End if (opf.ShowDialog() == DialogResult.OK)
        }

        private void toggleGuiState(bool isEnabled)
        {
            btnChangeCustomerDataFilePath.Enabled = isEnabled;
            btnConvert.Enabled = isEnabled;
            btnOpenRtf.Enabled = isEnabled;
            btnOpenXml.Enabled = isEnabled;
        }
    }
}
