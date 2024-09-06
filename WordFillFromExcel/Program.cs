using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml; // EPPlus library
using Spire.Doc;
using Document = Spire.Doc.Document;

class Program
{
    static void Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.InputEncoding = System.Text.Encoding.UTF8;

        // Ask user for input paths and file names
        Console.WriteLine("Enter the full path to the Word document template (e.g., C:\\path\\to\\template.docx):");
        string templateDocPath = Console.ReadLine();

        Console.WriteLine("Enter the full path to the Excel file (e.g., C:\\path\\to\\data.xlsx):");
        string excelPath = Console.ReadLine();

        Console.WriteLine("Enter the output folder path where the PDF files will be saved (e.g., C:\\path\\to\\output\\folder):");
        string outputFolder = Console.ReadLine();
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        Console.WriteLine("Enter the prefix for the document names (e.g., استدعاء_):");
        string docPrefix = Console.ReadLine();

        // Ask user for the placeholders in the Word document to be replaced
        var replacements = new Dictionary<string, List<int>>();
        string placeholder;
        do
        {
            Console.WriteLine("Enter the placeholder text in the Word document to replace (or 'done' to finish):");
            placeholder = Console.ReadLine();

            if (placeholder.ToLower() != "done")
            {
                Console.WriteLine($"Enter the column indices (as integer values) from the Excel file to replace '{placeholder}' (comma-separated if multiple columns):");
                string columnsInput = Console.ReadLine();
                var columns = columnsInput.Split(',').Select(int.Parse).ToList();
                replacements.Add(placeholder, columns);
            }

        } while (placeholder.ToLower() != "done");

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Load the Excel data
        using (var package = new ExcelPackage(new FileInfo(excelPath)))
        {
            var worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first worksheet
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++) // Start from row 2 to skip the header
            {
                // Collect values for the document name (e.g., prenomFr, nomFr) from Excel columns
                string prenomFr = worksheet.Cells[row, 2].Text.Trim();  // "prenom_fr" in column B
                string nomFr = worksheet.Cells[row, 3].Text.Trim();     // "nom_fr" in column C

                // Construct the new file name
                string newDocName = $"{docPrefix}{prenomFr}_{nomFr}.docx";
                string newDocPath = Path.Combine(outputFolder, newDocName);

                // Copy the template document to the new path
                File.Copy(templateDocPath, newDocPath, true);

                // Open and modify the new Word document
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(newDocPath, true))
                {
                    var body = wordDoc.MainDocumentPart.Document.Body;

                    // Loop through each replacement
                    foreach (var replacement in replacements)
                    {
                        string placeholderText = replacement.Key;
                        List<int> columns = replacement.Value;

                        // Build the replacement text from the specified Excel columns
                        var replacementText = string.Join(" ", columns.Select(col => worksheet.Cells[row, col].Text.Trim()));

                        // Replace the placeholder in the Word document
                        foreach (var run in body.Descendants<Run>())
                        {
                            var textElement = run.Elements<Text>().FirstOrDefault();
                            if (textElement != null && textElement.Text.Contains(placeholderText))
                            {
                                textElement.Text = textElement.Text.Replace(placeholderText, replacementText);
                            }
                        }
                    }
                }

                // Convert Word document to PDF
                string pdfFilePath = Path.Combine(outputFolder, $"{docPrefix}{prenomFr}_{nomFr}.pdf");
                ConvertWordToPdf(newDocPath, pdfFilePath);
                Console.WriteLine($"Generated PDF: {pdfFilePath}");
            }
        }

        Console.WriteLine("All documents created and saved as PDFs.");
    }

    static void ConvertWordToPdf(string wordFilePath, string pdfFilePath)
    {
        // Load the Word document
        Document document = new Document();
        document.LoadFromFile(wordFilePath);

        // Save as PDF
        document.SaveToFile(pdfFilePath, FileFormat.PDF);
        Console.WriteLine($"Converted Word to PDF: {pdfFilePath}");
    }
}
