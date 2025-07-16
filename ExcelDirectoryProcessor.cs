using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ExcelEncountersAutofill
{
    internal class ExcelDirectoryProcessor
    {
        private readonly List<NameMapping> _mappings;

        internal ExcelDirectoryProcessor(List<NameMapping> mappings)
        {
            _mappings = mappings;
        }

        internal void ProcessDirectory()
        {
            // Get the directory where the exe is being executed.
            string currentDirectory = Directory.GetCurrentDirectory();

            string[] excelFiles = Directory.GetFiles(currentDirectory, "*.xlsx");

            if (excelFiles.Length == 0)
            {
                Console.WriteLine("No Excel files found in the directory: " + currentDirectory);
            }

            foreach (var file in excelFiles)
            {
                try
                {
                    ProcessFile(file);
                    Console.WriteLine($"Processed file: {Path.GetFileName(file)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file {Path.GetFileName(file)}: {ex.Message}");
                }
            }
        }

        private void ProcessFile(string filePath)
        {
            using var package = new ExcelPackage(new FileInfo(filePath));
            if (package.Workbook.Worksheets.Count < 1)
            {
                throw new Exception("No worksheets found in the Excel file.");
            }

            // Select the last worksheet available
            var worksheet = package.Workbook.Worksheets[^1];

            int headerRow = 1;
            int startRow = headerRow + 1; // Data starts at row 2 (after headers)
            int nameColumn = 1; // Names are in column A
            int spokenAsColumn = 2; // Spoken version will be written in column B
            int totalRows = worksheet.Dimension.End.Row;
            int consecutiveEmptyRows = 0;

            // Process rows starting from the second row
            for (int row = startRow; row <= totalRows; row++)
            {
                string cellValue = worksheet.Cells[row, nameColumn].Text.Trim();

                // Count empty rows to determine stop condition.
                if (string.IsNullOrEmpty(cellValue))
                {
                    consecutiveEmptyRows++;
                    if (consecutiveEmptyRows > 3)
                    {
                        // Stop if more than three consecutive empty rows encountered.
                        break;
                    }
                    continue;
                }

                consecutiveEmptyRows = 0;

                // Look for a matching mapping (case-insensitive)
                var mapping = _mappings.FirstOrDefault(m => m.Names.Any(name =>
                    string.Equals(name.Trim(), cellValue, StringComparison.InvariantCultureIgnoreCase)));

                if (mapping != null)
                {
                    // Write the spoken version into the adjacent column (column B)
                    worksheet.Cells[row, spokenAsColumn].Value = mapping.SpokenAs;
                }
            }

            // Save any changes back to the Excel file.
            package.Save();
        }
    }
}