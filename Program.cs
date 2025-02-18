using System.Collections.Generic;
using System;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Threading;
using HtmlAgilityPack;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Tesseract;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Xml.Linq;
using DocumentFormat.OpenXml;


namespace wordToHtmlConverter
{
    class Program
    {
        public static List<Option> functionalityOptions;
        public static List<Option> wordOptions;
        public static List<Option> htmlOptions;
        public static List<Option> pdfOptions;

        static void Main(string[] args)
        {
            int extensionIndex = 0;

            htmlOptions = new List<Option>
            {
                new Option("Quitar formatos a documento .html", () => cleanHtml("*.html")),
                new Option("Convertir html a plantilla de outlook oft", () => htmlToOft("*.html")),
                new Option("Salir", () => writePrincipalMenu()),
            };

            wordOptions = new List<Option>
            {
                new Option("Convertir .doc a html", () => convertDocuments("*.doc")),
                new Option("Convertir .docx a html", () =>  convertDocuments("*.docx")),
                new Option("Remover headers y footers", () => RemoveHeadersAndFooters("*.docx")),
                new Option("Salir", () => writePrincipalMenu()),
            };

            pdfOptions = new List<Option>
            {
                new Option("Convertir pdf a html", () => convertPdfToHtml("*.pdf")),
                new Option("Eliminar todas las imagenes a PDF", () => removePdfImages("*.pdf")),
                new Option("Salir", () => writePrincipalMenu()),
            };

            // Create options that you want your menu to have
            functionalityOptions = new List<Option>
            {
                new Option("Html", () => WriteMenu(htmlOptions, htmlOptions[extensionIndex], "Seleccione:")),
                new Option("Word", () =>  WriteMenu(wordOptions, wordOptions[extensionIndex], "Seleccione:")),
                new Option("Pdf", () =>  WriteMenu(pdfOptions, wordOptions[extensionIndex], "Seleccione:")),
                new Option("Salir", () => Environment.Exit(0)),
            };
            writePrincipalMenu();
            //WriteMenu(options, options[index]);
            Console.ReadKey();
        }

        static void writePrincipalMenu() {
            int index = 0;
            WriteMenu(functionalityOptions, functionalityOptions[index], "Con que tipo de documentos desea trabajar?");
        }

        static void moveMenu(List<Option> options, int index, string instruction) {
            // Store key info in here
            ConsoleKeyInfo keyinfo;
            do
            {
                keyinfo = Console.ReadKey();

                // Handle each key input (down arrow will write the menu again with a different selected item)
                if (keyinfo.Key == ConsoleKey.DownArrow)
                {
                    if (index + 1 < options.Count)
                    {
                        index++;
                        WriteMenu(options, options[index], instruction);
                    }
                }
                if (keyinfo.Key == ConsoleKey.UpArrow)
                {
                    if (index - 1 >= 0)
                    {
                        index--;
                        WriteMenu(options, options[index], instruction);
                    }
                }
                // Handle different action for the option
                if (keyinfo.Key == ConsoleKey.Enter)
                {
                    options[index].Selected.Invoke();
                    index = 0;
                }
            }
            while (keyinfo.Key != ConsoleKey.X);
        }

        static void WriteMenu(List<Option> options, Option selectedOption, string instruction)
        {
            Console.Clear();

            Console.WriteLine(instruction);

            foreach (Option option in options)
            {
                if (option == selectedOption)
                {
                    Console.Write("> ");
                }
                else
                {
                    Console.Write(" ");
                }

                Console.WriteLine(option.Name);
            }

            moveMenu(options, options.FindIndex(x => x == selectedOption), instruction);
        }

        static void convertDocuments(string extension)
        {

            Console.WriteLine("Digite la ruta de los archivos");
            string directoryPath = Console.ReadLine();

            Console.WriteLine("Digite la ruta donde desea guardar los html");
            string htmlFolderPath = Console.ReadLine();

            // Get all word files in the specified directory
            string[] wordFiles = Directory.GetFiles(directoryPath, extension);

            // Create a Word application object
            Application wordApp = new Application();

            // Disable alerts and visible Word application
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            wordApp.Visible = false;

            foreach (string wordFile in wordFiles)
            {
                try
                {

                    Console.WriteLine("Converting: " + wordFile);

                    // Open the Word document
                    Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(wordFile);

                    // Generate a unique file name for the HTML file
                    string htmlFileName = System.IO.Path.GetFileNameWithoutExtension(wordFile) + ".html";
                    string htmlFilePath = System.IO.Path.Combine(htmlFolderPath, htmlFileName);

                    // Save the Word document as HTML
                    doc.SaveAs2(htmlFilePath, WdSaveFormat.wdFormatFilteredHTML);

                    // Close the Word document
                    doc.Close();

                    //RemoveHeadTagFromHtml(htmlFilePath);

                    Console.WriteLine($"Converted '{wordFile}' to '{htmlFilePath}'");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{wordFile}': {ex.Message}");
                }
            }

            // Close the Word application
            wordApp.Quit();

            // Release the COM objects
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);

            Console.WriteLine("Conversion completeda.");

            Thread.Sleep(3000);
            writePrincipalMenu();
            //WriteMenu(options, options.First());
        }

        static void cleanHtml(string extension) {
            Console.WriteLine("Digite la ruta de los archivos");
            string directoryPath = Console.ReadLine();

            // Get all html files in the specified directory
            string[] htmlFiles = Directory.GetFiles(directoryPath, extension);

            foreach (string htmlFile in htmlFiles) {
                HtmlDocument doc = new HtmlDocument();
                doc.Load(htmlFile);

                // Eliminar etiquetas <style>
                var styleNodes = doc.DocumentNode.SelectNodes("//style");
                if (styleNodes != null)
                {
                    foreach (var node in styleNodes)
                    {
                        node.Remove();
                    }
                }

                // Eliminar atributos de estilo en línea
                var nodesWithStyleAttribute = doc.DocumentNode.SelectNodes("//*[@style]");
                if (nodesWithStyleAttribute != null)
                {
                    foreach (var node in nodesWithStyleAttribute)
                    {
                        node.Attributes["style"].Remove();
                    }
                }

                // Eliminar clases CSS
                var nodesWithClassAttribute = doc.DocumentNode.SelectNodes("//*[@class]");
                if (nodesWithClassAttribute != null)
                {
                    foreach (var node in nodesWithClassAttribute)
                    {
                        node.Attributes["class"].Remove();
                    }
                }

                // Eliminar atributo align
                var nodesWithAlignAttribute = doc.DocumentNode.SelectNodes("//*[@align]");
                if (nodesWithAlignAttribute != null)
                {
                    foreach (var node in nodesWithAlignAttribute)
                    {
                        node.Attributes["align"].Remove();
                    }
                }

                //Eliminar elementos html vacíos
                RemoveEmptyNodes(doc.DocumentNode);

                //Verificar si hay elementos iguales seguidos y fusionarlos en uno solo
                MergeAdjacentEqualTags(doc.DocumentNode);

                // Guardar el documento modificado
                doc.Save(htmlFile);
            }
            Console.WriteLine("Todos los estilos y formatos han sido eliminados.");

            Thread.Sleep(3000);
            writePrincipalMenu();
        }
        
        static void htmlToOft(string extension)
        {
            Console.Write("Ingrese la ruta del archivo HTML: ");
            string rutaHtml = Console.ReadLine();

            // Verificar si la ruta es válida
            if (System.IO.File.Exists(rutaHtml))
            {
                Console.WriteLine("Archivo encontrado. Procesando...");
                // Llamar a tu método para procesar el archivo
                var outlookApp = new Microsoft.Office.Interop.Outlook.Application();
                var mailItem = (Microsoft.Office.Interop.Outlook.MailItem)outlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                // Leer el contenido HTML
                string htmlBody = System.IO.File.ReadAllText(rutaHtml);
                mailItem.HTMLBody = htmlBody;

                // Guardar como OFT
                // Obtener el nombre base del archivo y cambiar la extensión a .oft
                string nombreArchivoOft = System.IO.Path.ChangeExtension(rutaHtml, ".oft");
                mailItem.SaveAs(nombreArchivoOft, Microsoft.Office.Interop.Outlook.OlSaveAsType.olTemplate);

                Console.WriteLine($"Archivo OFT guardado como: {nombreArchivoOft}");
                Thread.Sleep(1000);
                writePrincipalMenu();
            }
            else
            {
                Console.WriteLine("La ruta del archivo no es válida o el archivo no existe.");
            }
        }
            static void RemoveEmptyNodes(HtmlNode node)
        {
            if (node == null)
                return;

            var emptyNodes = node.Descendants()
                                 .Where(n => !n.HasChildNodes && string.IsNullOrWhiteSpace(n.InnerText))
                                 .ToList();

            foreach (var emptyNode in emptyNodes)
            {
                emptyNode.Remove();
            }

            // Eliminar nodos que se vuelven vacíos después de eliminar sus hijos vacíos
            var nodesWithChildren = node.Descendants().Where(n => n.HasChildNodes).ToList();
            foreach (var parentNode in nodesWithChildren)
            {
                RemoveEmptyNodes(parentNode);
            }

            if (!node.HasChildNodes && string.IsNullOrWhiteSpace(node.InnerText))
            {
                node.Remove();
            }
        }

        static void MergeAdjacentEqualTags(HtmlNode node)
        {
            if (node == null)
                return;

            foreach (var childNode in node.ChildNodes.ToList())
            {
                MergeAdjacentEqualTags(childNode);
            }

            for (int i = 0; i < node.ChildNodes.Count - 1; i++)
            {
                var current = node.ChildNodes[i];
                var next = node.ChildNodes[i + 1];

                if (current.Name == next.Name && current.Name != "#text")
                {
                    // Fusionar el contenido del siguiente nodo en el nodo actual
                    current.InnerHtml +=  next.InnerHtml + "<br>";

                    // Eliminar el nodo siguiente
                    next.Remove();

                    // Retrocede un paso para verificar la nueva combinación con el siguiente nodo
                    i--;
                }
            }
        }

        static void RemoveHeadTagFromHtml(string filePath)
        {
            string htmlContent = File.ReadAllText(filePath);
            int headStartIndex = htmlContent.IndexOf("<head>", StringComparison.OrdinalIgnoreCase);
            int headEndIndex = htmlContent.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);

            if (headStartIndex >= 0 && headEndIndex >= 0)
            {
                headEndIndex += "</head>".Length;
                htmlContent = htmlContent.Remove(headStartIndex, headEndIndex - headStartIndex);
                File.WriteAllText(filePath, htmlContent);
            }
        }

        static void RemoveHeadersAndFooters(string extension)
        {
            // Specify the folder containing the Word documents
            Console.WriteLine("Digite la ruta de los archivos");
            string documentsFolder = Console.ReadLine();

            // Get a list of Word documents in the folder
            string[] docFiles = Directory.GetFiles(documentsFolder, extension);

            foreach (string docFile in docFiles)
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(docFile, true))
                {
                    // Get the MainDocumentPart
                    MainDocumentPart mainPart = doc.MainDocumentPart;

                    // Get the headers and footers parts
                    var headerParts = mainPart.HeaderParts;
                    var footerParts = mainPart.FooterParts;

                    // Remove the headers
                    foreach (var headerPart in headerParts)
                    {
                        headerPart.Header.RemoveAllChildren();
                        headerPart.Header.Save();
                    }

                    // Remove the footers
                    foreach (var footerPart in footerParts)
                    {
                        footerPart.Footer.RemoveAllChildren();
                        footerPart.Footer.Save();
                    }
                }
            }

            Console.WriteLine("Tarea completeda.");

            Thread.Sleep(3000);
            writePrincipalMenu();
        }

        static void convertPdfToHtml(string extension) {
            Console.WriteLine("Digite la ruta de los archivos pdf");
            string directoryPath = Console.ReadLine();

            Console.WriteLine("Digite la ruta donde desea guardar los html");
            string htmlPath = Console.ReadLine();

            // Get all word files in the specified directory
            string[] pdfFiles = Directory.GetFiles(directoryPath, extension);


            foreach (string pdfPath in pdfFiles)
            {
                try
                {

                 // Generate a unique file name for the HTML file
                 string htmlFileName = System.IO.Path.GetFileNameWithoutExtension(pdfPath) + ".html";
                 string htmlFilePath = System.IO.Path.Combine(htmlPath, htmlFileName);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{pdfPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Conversion completeda.");

            Thread.Sleep(3000);
            writePrincipalMenu();
        }

        static void removePdfImages(string extension)
        {
            Console.WriteLine("Digite la ruta de los archivos pdf");
            string directoryPath = Console.ReadLine();

            //Console.WriteLine("Digite la ruta donde desea guardar los nuevos pdf");
            //string pdfSavePath = Console.ReadLine();

            // Get all word files in the specified directory
            string[] pdfFiles = Directory.GetFiles(directoryPath, extension);

            foreach (string pdfPath in pdfFiles)
            {
                try
                {
                    string outputPdfPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(pdfPath), "Processed_" + System.IO.Path.GetFileName(pdfPath));
                    //string outputPdfPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(pdfSavePath), System.IO.Path.GetFileName(pdfPath));

                    using (PdfReader reader = new PdfReader(pdfPath))
                    {
                        using (PdfStamper stamper = new PdfStamper(reader, new FileStream(outputPdfPath, FileMode.Create)))
                        {
                            for (int i = 1; i <= reader.NumberOfPages; i++)
                            {
                                PdfDictionary pageDict = reader.GetPageN(i);
                                PdfObject obj = pageDict.GetDirectObject(PdfName.RESOURCES);
                                PdfDictionary resourcesDict = (PdfDictionary)PdfReader.GetPdfObject(obj);

                                if (resourcesDict != null)
                                {
                                    PdfDictionary xObjectDict = resourcesDict.GetAsDict(PdfName.XOBJECT);

                                    if (xObjectDict != null)
                                    {

                                        foreach (PdfName name in xObjectDict.Keys)
                                        {
                                            PdfObject xObject = xObjectDict.Get(name);

                                            if (xObject.IsIndirect())
                                            {
                                                PdfDictionary dict = (PdfDictionary)PdfReader.GetPdfObject(xObject);

                                                PdfName subtype = dict.GetAsName(PdfName.SUBTYPE);

                                                if (subtype != null && subtype.Equals(PdfName.IMAGE))
                                                {
                                                    xObjectDict.Remove(name); // Elimina la imagen
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{pdfPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Conversion completeda.");
            Console.ReadLine();
            Thread.Sleep(3000);
            writePrincipalMenu();
        }

    }

    public class Option
    {
        public string Name { get; }
        public Action Selected { get; }

        public Option(string name, Action selected)
        {
            Name = name;
            Selected = selected;
        }
    }
}
