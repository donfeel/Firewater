using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Printing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;
using System.Xml;

namespace Firewater.Viagra
{
    public class XpsCreate
    {
        const string packageWithPrintTicketName = @"C:\Users\Dhe\Desktop\xps\XpsDocumentithPitTicket.xps";
        const string packageName = @"C:\Users\Dhe\Desktop\xps\XpsDocument.xps"; // (without PrintTicket)
        const string backImg = @"C:\Users\Dhe\Desktop\xps\content\template.jpg";
        const string image1 = @"C:\Users\Dhe\Desktop\xps\content\picture.jpg";
        const string image2 = @"C:\Users\Dhe\Desktop\xps\content\image.tif";
        const string font1 = @"C:\Users\Dhe\Desktop\xps\content\courier.ttf";
        const string font2 = @"C:\Users\Dhe\Desktop\xps\content\msyh.ttf";
        const string fontContentType =
            "application/vnd.ms-package.obfuscated-opentype";

        const string PageContent_X = "0";
        const string PageContent_Y = "0";

        const int CONTENT_CUSTOMCODE_X = 98;
        const string CONTENT_CUSTOMCODE_Y = "152";
        const double CONTENT_CUSTOMCODE_SHIFT = 22.5;

        const string CONTENT_SENDERCOMPANY_X = "90";
        const string CONTENT_SENDERCOMPANY_Y = "186";

        const string CONTENT_SENDERNAME_X = "332";
        const string CONTENT_SENDERNAME_Y = "186";

        const string CONTENT_SENDERADDRESS_X = "90";
        const string CONTENT_SENDERADDRESS_Y = "218";

        private string[] customCODE = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D" };//客户编码
        private string senderCompany = "美库尔(上海)咨询公司"; //寄件公司
        private string senderName = "超人"; // 联络人 
        private string senderAddress = "上海市浦东新区张扬路3611弄A座1101室"; // 寄件人地址

        // -------------------------------- Run -----------------------------------
        /// <summary>
        ///   Creates two XpsDocument packages, the first without a PrintTicket
        ///   and a second with a PrintTicket.</summary>
        public void Run(string value)
        {
            //content = value;
            // First, create an XpsDocument without a PrintTicket. - - - - - - - -

            // If the document package exists from a previous run, delete it.
            if (File.Exists(packageName))
                File.Delete(packageName);

            //<SnippetXpsCreatePackageOpen>
            // Create an XpsDocument package (without PrintTicket).
            //using (Package package = Package.Open(packageName))
            //{
            //    XpsDocument xpsDocument = new XpsDocument(package);

            //    // Add the package content (false=without PrintTicket).
            //    AddPackageContent(xpsDocument, false);

            //    // Close the package.
            //    xpsDocument.Close();
            //}
            //</SnippetXpsCreatePackageOpen>

            // Next, create a second XpsDocument with a PrintTicket. - - - - - - -

            // If the document package exists from a previous run, delete it.
            if (File.Exists(packageWithPrintTicketName))
                File.Delete(packageWithPrintTicketName);

            // Create an XpsDocument with PrintTicket.
            using (Package package2 = Package.Open(packageWithPrintTicketName))
            {
                XpsDocument xpsDocumentWithPrintTicket = new XpsDocument(package2);

                // Add the package content (true=with PrintTicket).
                AddPackageContent(xpsDocumentWithPrintTicket, true);

                // Close the package.
                xpsDocumentWithPrintTicket.Close();
            }

            // Normal Completion, show the package names created.
            string msg = "Created two XPS document packages:\n   - " +
                packageName + "\n   - " + packageWithPrintTicketName;

        }// end:Run()


        //<SnippetXpsCreateAddPkgContent>
        // ------------------------- AddPackageContent ----------------------------
        /// <summary>
        ///   Adds a predefined set of content to a given XPS document.</summary>
        /// <param name="xpsDocument">
        ///   The package to add the document content to.</param>
        /// <param name="attachPrintTicket">
        ///   true to include a PrintTicket with the
        ///   document; otherwise, false.</param>
        private void AddPackageContent(
            XpsDocument xpsDocument, bool attachPrintTicket)
        {
            try
            {
                PrintTicket printTicket = GetPrintTicketFromPrinter();
                // PrintTicket is null, there is no need to attach one.
                if (printTicket == null)
                    attachPrintTicket = false;

                // Add a FixedDocumentSequence at the Package root
                IXpsFixedDocumentSequenceWriter documentSequenceWriter =
                    xpsDocument.AddFixedDocumentSequence();

                // Add the 1st FixedDocument to the FixedDocumentSequence. - - - - -
                IXpsFixedDocumentWriter fixedDocumentWriter =
                    documentSequenceWriter.AddFixedDocument();

                // Add content to the 1st document
                AddDocumentContent(fixedDocumentWriter);

                // Commit the 1st Document
                fixedDocumentWriter.Commit();

                // If attaching PrintTickets, attach one at
                // the package FixedDocumentSequence level.
                if (attachPrintTicket)
                    documentSequenceWriter.PrintTicket = printTicket;

                // Commit the FixedDocumentSequence
                documentSequenceWriter.Commit();
            }
            catch (XpsPackagingException xpsException)
            {
                throw xpsException;
            }
        }// end:AddPackageContent()
        //</SnippetXpsCreateAddPkgContent>


        //<SnippetXpsCreateAddDocContent>
        // ------------------------- AddDocumentContent ---------------------------
        /// <summary>
        ///   Adds a predefined set of content to a given document writer.</summary>
        /// <param name="fixedDocumentWriter">
        ///   The document writer to add the content to.</param>
        private void AddDocumentContent(IXpsFixedDocumentWriter fixedDocumentWriter)
        {
            // Collection of image and font resources used on the current page.
            //   Key: "XpsImage", "XpsFont"
            //   Value: List of XpsImage or XpsFont resources
            Dictionary<string, List<XpsResource>> resources;

            try
            {
                // Add Page 1 to current document.
                IXpsFixedPageWriter fixedPageWriter =
                    fixedDocumentWriter.AddFixedPage();

                // Add the resources for Page 1 and get the resource collection.
                resources = AddPageResources(fixedPageWriter);

                // Write page content for Page 1.
                WritePageContent(fixedPageWriter.XmlWriter,
                    "Page 1 of " + fixedDocumentWriter.Uri.ToString(), resources);

                // Commit Page 1.
                fixedPageWriter.Commit(); 
                
            }
            catch (XpsPackagingException xpsException)
            {
                throw xpsException;
            }
        }// end:AddDocumentContent()
        //</SnippetXpsCreateAddDocContent>


        //<SnippetXpsCreateAddPageResources>
        // -------------------------- AddPageResources ----------------------------
        private Dictionary<string, List<XpsResource>>
                AddPageResources(IXpsFixedPageWriter fixedPageWriter)
        {
            // Collection of all resources for this page.
            //   Key: "XpsImage", "XpsFont"
            //   Value: List of XpsImage or XpsFont
            Dictionary<string, List<XpsResource>> resources =
                new Dictionary<string, List<XpsResource>>();

            // Collections of images and fonts used in the current page.
            List<XpsResource> xpsImages = new List<XpsResource>();
            List<XpsResource> xpsFonts = new List<XpsResource>();

            try
            {
                XpsImage xpsImage;
                XpsFont xpsFont;

                // Add, Write, and Commit image1 to the current page.
                xpsImage = fixedPageWriter.AddImage(XpsImageType.JpegImageType);
                WriteToStream(xpsImage.GetStream(), backImg);
                xpsImage.Commit();
                xpsImages.Add(xpsImage);    // Add image1 as a required resource.

                // Add, Write, and Commit font 1 to the current page.
                xpsFont = fixedPageWriter.AddFont();
                WriteObfuscatedStream(
                    xpsFont.Uri.ToString(), xpsFont.GetStream(), font1);
                xpsFont.Commit();
                xpsFonts.Add(xpsFont);      // Add font1 as a required resource.

                // Add, Write, and Commit image2 to the current page.
                xpsImage = fixedPageWriter.AddImage(XpsImageType.TiffImageType);
                WriteToStream(xpsImage.GetStream(), image2);
                xpsImage.Commit();
                xpsImages.Add(xpsImage);    // Add image2 as a required resource.

                // Add, Write, and Commit font2 to the current page.
                xpsFont = fixedPageWriter.AddFont(false);
                WriteToStream(xpsFont.GetStream(), font2);
                xpsFont.Commit();
                xpsFonts.Add(xpsFont);      // Add font2 as a required resource.

                // Return the image and font resources in a combined collection.
                resources.Add("XpsImage", xpsImages);
                resources.Add("XpsFont", xpsFonts);
                return resources;
            }
            catch (XpsPackagingException xpsException)
            {
                throw xpsException;
            }
        }// end:AddPageResources()
        //</SnippetXpsCreateAddPageResources>


        //<SnippetPrinterCapabilities>
        // ---------------------- GetPrintTicketFromPrinter -----------------------
        /// <summary>
        ///   Returns a PrintTicket based on the current default printer.</summary>
        /// <returns>
        ///   A PrintTicket for the current local default printer.</returns>
        private PrintTicket GetPrintTicketFromPrinter()
        {
            PrintQueue printQueue = null;

            LocalPrintServer localPrintServer = new LocalPrintServer();

            // Retrieving collection of local printer on user machine
            PrintQueueCollection localPrinterCollection =
                localPrintServer.GetPrintQueues();

            System.Collections.IEnumerator localPrinterEnumerator =
                localPrinterCollection.GetEnumerator();

            if (localPrinterEnumerator.MoveNext())
            {
                // Get PrintQueue from first available printer
                printQueue = (PrintQueue)localPrinterEnumerator.Current;
            }
            else
            {
                // No printer exist, return null PrintTicket
                return null;
            }

            // Get default PrintTicket from printer
            PrintTicket printTicket = printQueue.DefaultPrintTicket;

            PrintCapabilities printCapabilites = printQueue.GetPrintCapabilities();

            // Modify PrintTicket
            if (printCapabilites.CollationCapability.Contains(Collation.Collated))
            {
                printTicket.Collation = Collation.Collated;
            }

            if (printCapabilites.DuplexingCapability.Contains(
                    Duplexing.TwoSidedLongEdge))
            {
                printTicket.Duplexing = Duplexing.TwoSidedLongEdge;
            }

            if (printCapabilites.StaplingCapability.Contains(Stapling.StapleDualLeft))
            {
                printTicket.Stapling = Stapling.StapleDualLeft;
            }

            return printTicket;
        }// end:GetPrintTicketFromPrinter()
        //</SnippetPrinterCapabilities>


        //<SnippetXpsCreateWritePageContent>
        // --------------------------- WritePageContent ---------------------------
        private void WritePageContent(XmlWriter xmlWriter, string documentUri,
            Dictionary<string, List<XpsResource>> resources)
        {
            List<XpsResource> xpsImages = resources["XpsImage"];
            List<XpsResource> xpsFonts = resources["XpsFont"];

            // Element are indented for reading purposes only
            xmlWriter.WriteStartElement("FixedPage");
            xmlWriter.WriteAttributeString("Width", "920");
            xmlWriter.WriteAttributeString("Height", "1056");
            xmlWriter.WriteAttributeString("xmlns",
                "http://schemas.microsoft.com/xps/2005/06");
            xmlWriter.WriteAttributeString("xml:lang", "en-US");

            xmlWriter.WriteStartElement("Path");
            xmlWriter.WriteAttributeString(
                "Data", "M 0,0 L 909,0 909,607 0,607 z");
            xmlWriter.WriteStartElement("Path.Fill");
            xmlWriter.WriteStartElement("ImageBrush");
            xmlWriter.WriteAttributeString(
                "ImageSource", xpsImages[0].Uri.ToString());
            xmlWriter.WriteAttributeString(
                "Viewbox", "0,0,909,607");
            //xmlWriter.WriteAttributeString("TileMode", "None");
            xmlWriter.WriteAttributeString(
                "ViewboxUnits", "Absolute");
            xmlWriter.WriteAttributeString(
                "ViewportUnits", "Absolute");
            xmlWriter.WriteAttributeString(
                "Viewport", "0,0,909,607");
            xmlWriter.WriteEndElement();
            xmlWriter.WriteEndElement();
            xmlWriter.WriteEndElement();

            // 客户编号
            var _shit = 0.0;
            foreach (var code in customCODE)
            {                
                xmlWriter.WriteStartElement("Glyphs");
                xmlWriter.WriteAttributeString("Fill", "#ff000000");
                xmlWriter.WriteAttributeString(
                    "FontUri", xpsFonts[0].Uri.ToString());
                xmlWriter.WriteAttributeString("FontRenderingEmSize", "18");
                xmlWriter.WriteAttributeString("OriginX", (CONTENT_CUSTOMCODE_X + _shit).ToString());
                xmlWriter.WriteAttributeString("OriginY", CONTENT_CUSTOMCODE_Y);
                xmlWriter.WriteAttributeString("UnicodeString", code);
                xmlWriter.WriteEndElement();
                _shit += CONTENT_CUSTOMCODE_SHIFT;
            }


            // 寄件公司
            xmlWriter.WriteStartElement("Glyphs");
            xmlWriter.WriteAttributeString("Fill", "#ff000000");
            xmlWriter.WriteAttributeString(
                "FontUri", xpsFonts[1].Uri.ToString());
            xmlWriter.WriteAttributeString("FontRenderingEmSize", "16");
            xmlWriter.WriteAttributeString("OriginX", CONTENT_SENDERCOMPANY_X);
            xmlWriter.WriteAttributeString("OriginY", CONTENT_SENDERCOMPANY_Y);
            //xmlWriter.WriteAttributeString(
            //    "UnicodeString", UnicodeToChinese(content));
            xmlWriter.WriteAttributeString(
                "UnicodeString", senderCompany);
            xmlWriter.WriteEndElement();

            //寄件人
            xmlWriter.WriteStartElement("Glyphs");
            xmlWriter.WriteAttributeString("Fill", "#ff000000");
            xmlWriter.WriteAttributeString(
                "FontUri", xpsFonts[1].Uri.ToString());
            xmlWriter.WriteAttributeString("FontRenderingEmSize", "16");
            xmlWriter.WriteAttributeString("OriginX", CONTENT_SENDERNAME_X);
            xmlWriter.WriteAttributeString("OriginY", CONTENT_SENDERNAME_Y);
            xmlWriter.WriteAttributeString(
                "UnicodeString", senderName);
            xmlWriter.WriteEndElement();

            //寄出地址
            xmlWriter.WriteStartElement("Glyphs");
            xmlWriter.WriteAttributeString("Fill", "#ff000000");
            xmlWriter.WriteAttributeString(
                "FontUri", xpsFonts[1].Uri.ToString());
            xmlWriter.WriteAttributeString("FontRenderingEmSize", "16");
            xmlWriter.WriteAttributeString("OriginX", CONTENT_SENDERADDRESS_X);
            xmlWriter.WriteAttributeString("OriginY", CONTENT_SENDERADDRESS_Y);
            xmlWriter.WriteAttributeString(
                "UnicodeString", senderAddress);
            xmlWriter.WriteEndElement();

            xmlWriter.WriteEndElement();
        }// end:WritePageContent()
        //<SnippetXpsCreateWritePageContent>


        // ----------------------------- WriteToStream ----------------------------
        private void WriteToStream(Stream stream, string resource)
        {
            const int bufSize = 0x1000;
            byte[] buf = new byte[bufSize];
            int bytesRead = 0;

            using (FileStream fileStream =
                new FileStream(resource, FileMode.Open, FileAccess.Read))
            {
                while ((bytesRead = fileStream.Read(buf, 0, bufSize)) > 0)
                {
                    stream.Write(buf, 0, bytesRead);
                }
            }
        }// end:WriteToStream()


        // ------------------------- WriteObfuscatedStream ------------------------
        private void WriteObfuscatedStream(
            string resourceName, Stream stream, string resource)
        {
            int bufSize = 0x1000;
            int guidByteSize = 16;
            int obfuscatedByte = 32;

            // Get the GUID byte from the resource name.  Typical Font name:
            //    /Resources/Fonts/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx.ODTTF
            int startPos = resourceName.LastIndexOf('/') + 1;
            int length = resourceName.LastIndexOf('.') - startPos;
            resourceName = resourceName.Substring(startPos, length);

            Guid guid = new Guid(resourceName);

            string guidString = guid.ToString("N");

            // Parsing the guid string and coverted into byte value
            byte[] guidBytes = new byte[guidByteSize];
            for (int i = 0; i < guidBytes.Length; i++)
            {
                guidBytes[i] = Convert.ToByte(guidString.Substring(i * 2, 2), 16);
            }

            using (FileStream filestream = new FileStream(resource, FileMode.Open))
            {
                // XOR the first 32 bytes of the source
                // resource stream with GUID byte.
                byte[] buf = new byte[obfuscatedByte];
                filestream.Read(buf, 0, obfuscatedByte);

                for (int i = 0; i < obfuscatedByte; i++)
                {
                    int guidBytesPos = guidBytes.Length - (i % guidBytes.Length) - 1;
                    buf[i] ^= guidBytes[guidBytesPos];
                }
                stream.Write(buf, 0, obfuscatedByte);

                // copy remaining stream from source without obfuscation
                buf = new byte[bufSize];

                int bytesRead = 0;
                while ((bytesRead = filestream.Read(buf, 0, bufSize)) > 0)
                {
                    stream.Write(buf, 0, bytesRead);
                }
            }
        }// end:WriteObfuscatedStream()


    }// end:class XpsCreate
}