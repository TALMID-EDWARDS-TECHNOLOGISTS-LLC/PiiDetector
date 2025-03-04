
using System.Collections.Generic;
using System.IO;
using System;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Linq;
using UglyToad.PdfPig;
using System.Runtime.CompilerServices;

namespace PiiDetector
{
    /// <summary>
    /// Detects Personally Identifiable Information (PII) in text or files using regex patterns.
    /// </summary>
    public class PiiDetector
    {
        private readonly List<Regex> patterns;

        /// <summary>
        /// Initializes a new instance of the <see cref="PiiDetector"/> class with predefined regex patterns for PII detection.
        /// </summary>
        public PiiDetector()
        {
            patterns = new List<Regex>
            {
                // Universal patterns
                new Regex(@"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", RegexOptions.Compiled), // Email
                new Regex(@"\b(\+?\d{1,3}[-.\s]?)?(\(?\d{3}\)?[-.\s]?)?\d{3}[-.\s]?\d{4}\b", RegexOptions.Compiled), // Phone
                new Regex(@"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b", RegexOptions.Compiled), // IPv4
                new Regex(@"([0-9a-fA-F]{1,4}:){7}[0-9a-fA-F]{1,4}", RegexOptions.Compiled), // IPv6
                new Regex(@"\b\d{1,2}/\d{1,2}/\d{4}\b|\b\d{4}-\d{2}-\d{2}\b", RegexOptions.Compiled), // Date

                // US-specific
                new Regex(@"\b\d{3}-\d{2}-\d{4}\b", RegexOptions.Compiled), // SSN
                new Regex(@"\b\d{9}\b", RegexOptions.Compiled), // Passport
                new Regex(@"\b[A-Z]?\d{8,12}\b", RegexOptions.Compiled), // Driver’s License
                new Regex(@"\b\d{5}(-\d{4})?\b", RegexOptions.Compiled), // ZIP Code

                // UK-specific
                new Regex(@"\b[A-Z]{2}\d{6}[A-D]?\b", RegexOptions.Compiled), // NINO
                new Regex(@"\b\d{9}\b", RegexOptions.Compiled), // Passport
                new Regex(@"\b[A-Z]{1,2}\d[A-Z\d]?\s\d[A-Z]{2}\b", RegexOptions.Compiled), // Postcode
                new Regex(@"\b[A-Z]{1,2}\d{6,7}[A-Z]?\b", RegexOptions.Compiled), // Driver’s License

                // France-specific
                new Regex(@"\b[12]\s?\d{2}\s?\d{2}\s?\d{2}\s?\d{3}\s?\d{3}\b", RegexOptions.Compiled), // SSN
                new Regex(@"\b\d{2}[A-Z]{2}\d{5}\b", RegexOptions.Compiled), // Passport
                new Regex(@"\b0\d{1}\s?\d{2}\s?\d{2}\s?\d{2}\s?\d{2}\b", RegexOptions.Compiled), // Phone
                new Regex(@"\b\d{5}\b", RegexOptions.Compiled), // Postal Code

                // Canada-specific
                new Regex(@"\b\d{3}-\d{3}-\d{3}\b", RegexOptions.Compiled), // SIN
                new Regex(@"\b[A-Z]{2}\d{6}\b", RegexOptions.Compiled), // Passport
                new Regex(@"\b[A-Z]\d[A-Z]\s?\d[A-Z]\d\b", RegexOptions.Compiled), // Postal Code
                new Regex(@"\b[A-Z]\d{8,9}\b", RegexOptions.Compiled), // Driver’s License

                // Financial
                new Regex(@"\b4[0-9]{12}(?:[0-9]{3})?\b", RegexOptions.Compiled), // Visa
                new Regex(@"\b5[1-5][0-9]{14}\b", RegexOptions.Compiled), // MasterCard
                new Regex(@"\b[A-Z]{2}\d{2}[A-Z0-9]{4}\d{7}([A-Z0-9]?){0,16}\b", RegexOptions.Compiled), // IBAN

                // Additional
                new Regex(@"\b[A-Z][a-z]+\s[A-Z][a-z]+\b", RegexOptions.Compiled), // Names
                new Regex(@"\b\d+\s[A-Za-z]+\s[A-Za-z]+\b", RegexOptions.Compiled) // Addresses
            };
        }

        /// <summary>
        /// Determines whether the specified text contains any Personally Identifiable Information (PII).
        /// </summary>
        /// <param name="text">The text to check for PII.</param>
        /// <returns><c>true</c> if PII is detected; otherwise, <c>false</c>.</returns>
        public bool ContainsPii(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            return patterns.Exists(pattern => pattern.IsMatch(text));
        }

        /// <summary>
        /// Determines whether the file at the specified path contains any Personally Identifiable Information (PII).
        /// </summary>
        /// <param name="filePath">The path to the file to check.</param>
        /// <returns><c>true</c> if PII is detected in the file; otherwise, <c>false</c>.</returns>
        /// <exception cref="ArgumentException">Thrown when the file path is null, empty, or the file does not exist.</exception>
        /// <exception cref="NotSupportedException">Thrown when the file extension is not supported.</exception>
        /// <exception cref="InvalidOperationException">Thrown when text extraction from the file fails.</exception>
        public bool ContainsPiiFromFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                throw new ArgumentException("Invalid or non-existent file path.", nameof(filePath));

            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            string text;

            try
            {
                switch (extension)
                {
                    case ".txt":
                    case ".csv":
                    case ".vcf":    // vCard files
                    case ".ics":    // iCalendar files
                    case ".mht":    // MIME HTML files
                    case ".rtf":    // Rich Text Format files
                    case ".xml":
                        text = File.ReadAllText(filePath);
                        break;
                    case ".json":
                        text = ExtractTextFromJson(filePath);
                        break;
                    case ".xlsx":
                        text = ExtractTextFromXlsx(filePath);
                        break;
                    case ".pdf":
                        text = ExtractTextFromPdf(filePath);
                        break;
                    case ".docx":
                        text = ExtractTextFromDocx(filePath);
                        break;
                    default:
                        throw new NotSupportedException($"File extension '{extension}' is not supported.");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract text from file: {filePath}", ex);
            }

            return ContainsPii(text);
        }

        /// <summary>
        /// Adds a custom regex pattern to the list of PII detection patterns.
        /// </summary>
        /// <param name="pattern">The regex pattern to add.</param>
        /// <exception cref="ArgumentException">Thrown when the pattern is null, empty, or invalid.</exception>
        public void AddPattern(string pattern)
        {
            if (string.IsNullOrEmpty(pattern))
                throw new ArgumentException("Pattern cannot be null or empty.", nameof(pattern));

            try
            {
                var regex = new Regex(pattern, RegexOptions.Compiled);
                patterns.Add(regex);
            }
            catch (ArgumentException ex)
            {
                throw new ArgumentException("Invalid regex pattern.", nameof(pattern), ex);
            }
        }

        private string ExtractTextFromJson(string filePath)
        {
            try
            {
                var json = File.ReadAllText(filePath);
                var jsonObject = JsonConvert.DeserializeObject<JToken>(json);
                return ExtractTextFromJToken(jsonObject);
            }
            catch (JsonException ex)
            {
                throw new InvalidOperationException($"Failed to parse JSON file: {filePath}", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract text from JSON file: {filePath}", ex);
            }
        }

        private string ExtractTextFromJToken(JToken token)
        {
            if (token == null) return string.Empty;

            var text = new StringBuilder();
            if (token.Type == JTokenType.Object || token.Type == JTokenType.Array)
            {
                foreach (var child in token.Children())
                {
                    text.Append(ExtractTextFromJToken(child));
                }
            }
            else if (token.Type == JTokenType.String)
            {
                text.Append(token.ToString() + " ");
            }
            return text.ToString();
        }

        private string ExtractTextFromXlsx(string filePath)
        {
            try
            {
                var text = new StringBuilder();
                using (var doc = SpreadsheetDocument.Open(filePath, false))
                {
                    var workbookPart = doc.WorkbookPart;
                    if (workbookPart == null)
                        throw new InvalidOperationException("The Excel file does not contain a workbook part.");

                    foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                    {
                        if (sheet.Id != null && sheet.Id.HasValue)
                        {
                            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
                            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                            if (sheetData != null)
                            {
                                foreach (var row in sheetData.Elements<Row>())
                                {
                                    foreach (var cell in row.Elements<Cell>())
                                    {
                                        text.Append(GetCellValue(cell, workbookPart) + " ");
                                    }
                                }
                            }
                        }
                    }
                }
                return text.ToString();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract text from XLSX file: {filePath}", ex);
            }
        }

        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell == null || string.IsNullOrEmpty(cell.InnerText)) return string.Empty;

            var value = cell.InnerText;
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                if (int.TryParse(value, out int id))
                {
                    var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                    if (sharedStringTable != null && id >= 0 && id < sharedStringTable.ChildElements.Count)
                    {
                        value = sharedStringTable.ChildElements[id].InnerText ?? string.Empty;
                    }
                    else
                    {
                        value = string.Empty;
                    }
                }
                else
                {
                    value = string.Empty;
                }
            }
            return value;
        }

        private string ExtractTextFromPdf(string filePath)
        {
            try
            {
                var text = new StringBuilder();
                using (var document = PdfDocument.Open(filePath))
                {
                    foreach (var page in document.GetPages())
                    {
                        text.Append(page.Text);
                    }
                }
                return text.ToString();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract text from PDF file: {filePath}", ex);
            }
        }

        private string ExtractTextFromDocx(string filePath)
        {
            try
            {
                using (var doc = WordprocessingDocument.Open(filePath, false))
                {
                    if (doc.MainDocumentPart?.Document?.Body != null)
                    {
                        return doc.MainDocumentPart.Document.Body.InnerText;
                    }
                    else
                    {
                        throw new InvalidOperationException("The Word document does not contain a main document part or body.");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract text from DOCX file: {filePath}", ex);
            }
        }
    }
}