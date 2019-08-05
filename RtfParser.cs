using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using deliverynote2xml.tags;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Runtime.InteropServices;
using System_Windows_Forms = System.Windows.Forms;

namespace deliverynote2xml
{
    class RtfParser
    {
        public Dictionary<string, string> Company { get; private set; }
        public List<Dictionary<string, string>> Documents { get; private set; }
        public List<Dictionary<string, string>> Rows { get; private set; }
        private string Discounts { get; set; }

        public RtfParser(string path)
        {
            Application word = new Application();
            Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();
            Frames frames = null;

            object fileName = path;

            // Define an object to pass to the API for missing parameters
            object missing = System.Type.Missing;

            try
            {
                doc = word.Documents.Open(ref fileName,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing);

                this.Company = new Dictionary<string, string>();
                this.Documents = new List<Dictionary<string, string>>();
                this.Rows = new List<Dictionary<string, string>>();

                frames = doc.Frames;

                //Example delivery note "005.rtf":
                //
                //       <Company>
                //   4 -  <Name>FTORR Corp.</Name>
                //   5 -  <Address>5283 Laoreet St.</Address>
                //   6 -  <Zip>71272</Zip>
                //   6 -  <City>Bangor</City>
                //   6 -  <State>ME</State>
                //        <Country>United States</Country>
                //   15 - <FiscalCode>06392640964</FiscalCode>
                //        <VatCode>06392640964</VatCode>
                //       </Company>            
                //       <Documents>
                //         <Document>
                //   50 -    <CustomerCode>5476</CustomerCode>
                //   19 -    <SoldToName>Pellentesque Consulting</SoldToName>
                //   18 -    <ShipToName>Neque Et Associates</ShipToName>
                //   21 -    <ShipToAddress>Ap #327-6373 Fusce St.</ShipToAddress>
                //   23 -    <ShipToZip>47019</ShipToZip>
                //   23 -    <ShipToCity>Essex</ShipToCity>
                //   23 -    <ShipToState>VI</ShipToState>


                //   37 -    <Date>2019-01-22</Date> //Document date


                //   39 -    <Number>152</Number>
                //   54 -    <TransportReason>Sale</TransportReason>
                //   145 -    <TransportDateTime>22/01/2019 17:28</TransportDateTime>
                //           <Rows>
                //             <Row>
                //               <Code>2014</Code>
                //               <Description>Glipizide</Description>
                //               <Qty>12,00</Qty>
                //             </Row>
                //              ...
                //         </Document>
                //       </Documents>
                //
                // Note: not all fields are used.

                this.Company.Add("Name", frames[4].Range.Text);

                this.Company.Add("Address", frames[5].Range.Text);

                string[] companyLocationZipCityState = frames[6].Range.Text.Split(' ');

                this.Company.Add("Zip", companyLocationZipCityState[0]);

                StringBuilder City = new StringBuilder();

                for (int i = 1; i < companyLocationZipCityState.Length - 1; i++)
                {
                    City.Append(companyLocationZipCityState[i]);

                    if (i != companyLocationZipCityState.Length - 2)
                    {
                        City.Append(" ");
                    }
                }

                this.Company.Add("City", City.ToString());

                this.Company.Add("State", companyLocationZipCityState[companyLocationZipCityState.Length - 1]);

                this.Company.Add("Country", "United States");

                //this.Company.Add("FiscalCode", frames[15].Range.Text);

                //this.Company.Add("VatCode", frames[15].Range.Text);

                Dictionary<string, string> document = new Dictionary<string, string>();

                document.Add("CustomerCode", frames[50].Range.Text);

                ExcelSheetParser xlsParser
                    = new ExcelSheetParser(string.Format(@"{0}",
                        CodeManager.CustomerDataFilePath));

                this.Discounts = xlsParser.GetDiscounts(document["CustomerCode"]);

                document.Add("CustomerName", frames[19].Range.Text);

                document.Add("DeliveryName", frames[18].Range.Text);

                City = new StringBuilder();

                string[] deliveryLocationZipCityState = frames[23].Range.Text.Split(' ');

                for (int i = 1; i < deliveryLocationZipCityState.Length - 1; i++)
                {
                    City.Append(deliveryLocationZipCityState[i]);

                    if (i != deliveryLocationZipCityState.Length - 2)
                    {
                        City.Append(" ");
                    }
                }

                document.Add("DeliveryAddress", frames[21].Range.Text);

                document.Add("DeliveryZip", deliveryLocationZipCityState[0]);

                document.Add("DeliveryCity", City.ToString());

                document.Add("DeliveryState", deliveryLocationZipCityState[deliveryLocationZipCityState.Length - 1]);

                // document.Add("DocumentType", "D");

                string[] documentDateArray = frames[37].Range.Text.Split('/');
                Array.Reverse(documentDateArray);
                string documentDateString = String.Join("-", documentDateArray);

                document.Add("Date", documentDateString);

                document.Add("Number", (frames[39].Range.Text).Trim());

                //document1.Add("TransportReason", frames[54].Range.Text);

                //document1.Add("TransportDateTime", frames[145].Range.Text);

                this.Documents.Add(document);

                //Calculate number of pages
                Microsoft.Office.Interop.Word.WdStatistic numOfPagesStat
                    = Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages;

                int numOfPages = doc.ComputeStatistics(numOfPagesStat, ref missing);
                int currentPage = 1;

                int itemCodePos = 0;

                //Repeat code in the "while" block for each page
                while (currentPage <= numOfPages)
                {
                    //First item code (in each page) is always at the 62th position.
                    if (frames[62].Range.Text != null) //Check whether items exist
                    {

                        //In each page, the first item position can be found at multiple of 62
                        itemCodePos += 62;

                        int loopCounter = 0;

                        //frames[itemCodePos+4] is the "Qty" column of the first item. If this is empty, it means customer name and 
                        //address lines have been incorrectly used as the article description for one or more articles. 
                        while (frames[itemCodePos + 4].Range.Text == null)
                        {
                            //Let's explain this with an example: (see example file: "customer-data-in-description-field.rtf"
                            //
                            //if three lines of article description contain:
                            //
                            //   DESCRIPTION
                            //
                            //   Eu Corp.
                            //   3019 A, Av.
                            //   23807 Tulsa OK
                            //
                            //then: itemCodePos+3 is: "3019 A, Av."
                            //      itemCodePos+6 is: "23807 Tulsa OK"
                            //
                            //and so on incrementing by multiples of three
                            itemCodePos += 3;

                            //We know when at the next iteration we'll find the first item code, because at that
                            //point frames[itemCodePos + 4] is going to contain the first item description, hence it will
                            //not be empty anymore
                            if (frames[itemCodePos + 4].Range.Text != null)
                            {
                                //Again, referring to previous example, if frames[itemCodePos + 4] contains an item description,
                                //this means we are on the last line of the customer address (in the description column).
                                //
                                //If this is the last line of a customer address, then moving three positions onward, we'll find
                                //our first item code.
                                //
                                //So, here we move where the item code is and before exiting the loop.
                                itemCodePos += 3;

                                //Before exiting the loop, warn the user that the item description incorrectly contains the personal
                                //data of a customer
                                System_Windows_Forms.MessageBox.Show("Warning: the customer data in this delivery note has been incorrectly placed in the \"DESCRIPTION\" field of one or more items.\n\nThe XML file will be generated anyway.",
                                   "Document not well formed",
                                       System_Windows_Forms.MessageBoxButtons.OK,
                                           System_Windows_Forms.MessageBoxIcon.Warning);
                            }

                            //Just in case, if for any reason the code logic here turns into an infinite loop, exit
                            //after max 10000 iterations
                            if (loopCounter > 10000)
                            {
                                throw new Exception("error while fetching item data.");
                            }

                            loopCounter++;
                        }

                        int itemDescriptionPos = itemCodePos + 1;

                        //Oddly enough, quantity for the first item comes before its code
                        int itemQtyPos = itemCodePos - 1;

                        do
                        {
                            Dictionary<string, string> row = new Dictionary<string, string>();

                            string itemCode = frames[itemCodePos].Range.Text;

                            row.Add("Code", itemCode);
                            row.Add("Description", frames[itemDescriptionPos].Range.Text);
                            row.Add("Qty", frames[itemQtyPos].Range.Text);
                            row.Add("Discounts", this.Discounts);

                            this.Rows.Add(row);

                            if ((frames[itemDescriptionPos + 2].Range.Text != null) 
                                && (frames[itemDescriptionPos + 2].Range.Text.Trim().ToUpper() == "EXTERNAL PACKAGE"))
                            {

                                //itemDescriptionPos + 2 is the position of the frame containg the string: "EXTERNAL PACKAGE".
                                //Add 15 to it to move to the last frame in the page (it's in the footer on the right and i's empty)
                                //Remember, in each page the first item position can be found at multiple of 62
                                itemCodePos = itemDescriptionPos + 2 + 15;

                                //Used for testing:
                                //string testValue = frames[nextItemCodePos].Range.Text ?? "NULL";

                                break;
                            }

                            itemCodePos += 5;
                            itemDescriptionPos = itemCodePos + 1;
                            itemQtyPos = itemCodePos - 1;

                        } while (true);

                    }

                    currentPage++;
                }

            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                //Release com objects to fully kill Word process from running in 
                //the background
                CodeManager.ReleaseComObject(frames);
                frames = null;

                //Close and release
                ((_Document)doc).Close();
                CodeManager.ReleaseComObject(doc);
                doc = null;

                //Quit and release
                if (word != null)
                {
                    ((_Application)word).Quit();
                }

                CodeManager.ReleaseComObject(word);
                word = null;

                GC.Collect();
            }

        }
    }
}
