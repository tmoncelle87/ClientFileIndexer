using System;              // Basic C# types (Console, etc.)
using System.IO;           // File & directory access (Directory, Path, DirectoryInfo)
using System.Linq;         // LINQ helpers (OrderBy)
using ClosedXML.Excel;     // Excel (.xlsx) creation & formatting library

namespace FolderFileIndexer
{
    public class Program
    {

        public static void Main(string[] args)
        {
            // Folder we want to scan for subfolders
            string folderPath = @"C:\Users\tmonc\OneDrive - Minnesota State\Documents";

            // Output Excel file path
            string outputFile = @"C:\temp\folder_report.xlsx";

            // Ensure the output directory exists so SaveAs doesn't fail
            string? outputDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            string[] dirs;

            try
            {
                // Attempt to enumerate directories in the starting folder
                dirs = Directory.GetDirectories(folderPath);
            }
            catch (UnauthorizedAccessException)
            {
                Console.WriteLine("Access denied to starting folder.");
                return; // Exit Main safely
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read starting folder: {ex.Message}");
                return;
            }


            // Sort AFTER we successfully retrieved the directories
            var orderedDirs = dirs.OrderBy(d => d, StringComparer.OrdinalIgnoreCase);


            // Create a new Excel workbook in memory
            using (var workbook = new XLWorkbook())
            {
                // Add a worksheet named "Folders"
                var ws = workbook.Worksheets.Add("Folders");

                // Write column headers
                ws.Cell(1, 1).Value = "Index";
                ws.Cell(1, 2).Value = "Folder Name";

                // Format header row to be bold
                ws.Range(1, 1, 1, 2).Style.Font.Bold = true;

                int row = 2;     // Start writing data on row 2 (row 1 = headers)
                int count = 1;   // Running index counter

                // Loop through each directory path
                foreach (string dir in orderedDirs)

                {
                    // Column 1: numeric index
                    ws.Cell(row, 1).Value = count++;

                    // Column 2: folder name only (not full path)
                    ws.Cell(row, 2).Value = new DirectoryInfo(dir).Name;

                    row++; // Move to the next row
                }

                // Resize columns automatically based on content width
                ws.Columns().AdjustToContents();

                // Save the workbook to disk as a real .xlsx file
                workbook.SaveAs(outputFile);
            }

            // Console feedback so user knows it worked
            Console.WriteLine("Done. Excel file created: " + outputFile);
        }
    }
}
