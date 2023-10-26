using System.Collections.Generic;
using System;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Threading;
using System.Linq;

namespace wordToHtmlConverter
{
    class Program
    {
        public static List<Option> options;
        static void Main(string[] args)
        {


            // Create options that you want your menu to have
            options = new List<Option>
            {
                new Option(".doc", () => convertDocuments("*.doc")),
                new Option(".docx", () =>  convertDocuments("*.docx")),
                new Option("Salir", () => Environment.Exit(0)),
            };

            int index = 0;
            WriteMenu(options, options[index]);

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
                        WriteMenu(options, options[index]);
                    }
                }
                if (keyinfo.Key == ConsoleKey.UpArrow)
                {
                    if (index - 1 >= 0)
                    {
                        index--;
                        WriteMenu(options, options[index]);
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

            Console.ReadKey();
        }

        static void WriteMenu(List<Option> options, Option selectedOption)
        {
            Console.Clear();

            Console.WriteLine("Seleccione la extensión de los archivos");

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
                    Document doc = wordApp.Documents.Open(wordFile);

                    // Generate a unique file name for the HTML file
                    string htmlFileName = Path.GetFileNameWithoutExtension(wordFile) + ".html";
                    string htmlFilePath = Path.Combine(htmlFolderPath, htmlFileName);

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
            WriteMenu(options, options.First());


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
