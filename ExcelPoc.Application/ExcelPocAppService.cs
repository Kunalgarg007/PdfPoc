using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPoc.Contracts.DTO;
using ExcelPoc.Contracts.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelPoc.Application
{
    public class ExcelPocAppService : IExcelPocAppService
    {
        private readonly string _templatePath;
        private readonly string _savePath;

        public ExcelPocAppService()
        {
            _templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Templates", "NEW ALW ISP BLANK.xlsx");
            _savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Generated_Excels");
        }

        public void PrintAllNamedRanges(string filePath)
        {
            using var doc = SpreadsheetDocument.Open(filePath, false);
            var workbookPart = doc.WorkbookPart;

            Console.WriteLine("=== Named Ranges in Workbook ===");
            foreach (var name in workbookPart.Workbook.DefinedNames.Elements<DefinedName>())
            {
                Console.WriteLine($"{name.Name}: {name.Text}");
            }
        }

        public async Task<byte[]> GenerateExcelAsync(ExcelDto dto)
        {
            if (dto == null)
                throw new ArgumentNullException(nameof(dto));
            if (!File.Exists(_templatePath))
                throw new FileNotFoundException("Template not found", _templatePath);

            Directory.CreateDirectory(_savePath);
            string outputPath = Path.Combine(_savePath, $"Excel_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            File.Copy(_templatePath, outputPath, true);

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(outputPath, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart ?? throw new Exception("Workbook part missing");
                var definedNames = workbookPart.Workbook.DefinedNames?.Elements<DefinedName>().ToList() ?? new();

                // ✅ Define both named and direct cell mappings
                var fieldMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "rngFirstName", dto.Name ?? "" },
                    { "rngLastName", dto.LastName ?? "" },
                    { "E7", dto.OrganizationName ?? "" },          // direct cell
                    { "rngBirthdate", dto.BirthDate ?? "" },
                    { "D5", dto.PrintArea ?? "" },                 // direct cell
                    { "rngMedicare", dto.Medicare ?? "" },
                    { "rngMedications_TtlNum", dto.MedicationsTtlNum ?? "" }
                };

                foreach (var field in fieldMap)
                {
                    string fieldKey = field.Key;
                    string value = field.Value;

                    // Check if it's a direct cell reference (like "E7")
                    if (Regex.IsMatch(fieldKey, @"^[A-Z]+\d+$", RegexOptions.IgnoreCase))
                    {
                        Console.WriteLine($"Direct cell mapping detected: {fieldKey}");
                        // Use the first worksheet by default
                        var firstSheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                        if (firstSheet == null)
                        {
                            Console.WriteLine("No sheets found in workbook.");
                            continue;
                        }

                        WorksheetPart wsPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id);
                        SetCellValue(wsPart, fieldKey, value);
                        continue;
                    }

                    // Otherwise, treat as named range
                    var definedName = definedNames.FirstOrDefault(d => d.Name?.Value == fieldKey);
                    if (definedName == null)
                    {
                        Console.WriteLine($"Named range '{fieldKey}' not found in workbook.");
                        continue;
                    }

                    string rangeText = definedName.Text; // e.g. 'Sheet1!$F$11' or 'Sheet1!$F$11:$F$13'
                    string sheetName = rangeText.Split('!')[0].Trim('\'');
                    string firstCell = rangeText.Split('!')[1].Split(':')[0].Replace("$", "");

                    Sheet? sheet = workbookPart.Workbook.Descendants<Sheet>()
                        .FirstOrDefault(s => s.Name == sheetName);

                    if (sheet == null)
                    {
                        Console.WriteLine($"Sheet '{sheetName}' not found for range '{fieldKey}'.");
                        continue;
                    }

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SetCellValue(worksheetPart, firstCell, value);
                }

                workbookPart.Workbook.Save();
            }

            Console.WriteLine($"Editable Excel generated successfully: {outputPath}");
            return await File.ReadAllBytesAsync(outputPath);
        }

        private static void SetCellValue(WorksheetPart worksheetPart, string cellRef, string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                Console.WriteLine($"SheetData missing in worksheet for {cellRef}");
                return;
            }

            // Extract row number and column name
            uint rowIndex = uint.Parse(Regex.Replace(cellRef, "[^0-9]", ""));
            string columnName = new string(cellRef.Where(char.IsLetter).ToArray());

            // Find or create row
            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                row = new Row { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // Find or create cell
            Cell cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellRef);
            if (cell == null)
            {
                cell = new Cell { CellReference = cellRef };
                // maintain column order
                Cell? refCell = row.Elements<Cell>()
                    .FirstOrDefault(c => string.Compare(c.CellReference?.Value, cellRef, true) > 0);
                row.InsertBefore(cell, refCell);
            }

            // ✅ Set value properly
            cell.CellValue = new CellValue(value);
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
        }

        private static string GetCellValue(Cell cell, SharedStringTable? sharedStringTable)
        {
            if (cell.CellValue == null) return string.Empty;
            string value = cell.CellValue.InnerText;
            if (cell.DataType?.Value == CellValues.SharedString && sharedStringTable != null)
                return sharedStringTable.ElementAt(int.Parse(value)).InnerText;
            return value;
        }
    }
}
