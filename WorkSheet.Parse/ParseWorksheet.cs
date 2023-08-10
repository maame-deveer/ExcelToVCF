using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExtractAndConvert.Parse
{
    public class ParseWorkSheet
    {
        public List<WorkSheetRows> ExtractDataFromWorkSheet(string filePath)
        {
            List<WorkSheetRows> rows = new List<WorkSheetRows>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                foreach (Row row in sheetData.Elements<Row>())
                {
                    WorkSheetRows rowData = new WorkSheetRows();

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        string cellValue = GetCellValue(cell, workbookPart);
                        switch (GetCellColumnName(cell.CellReference))
                        {
                            case "A":
                                rowData.FirstName = cellValue;
                                break;
                            case "B":
                                rowData.LastName = cellValue;
                                break;
                            case "C":
                                rowData.Institution = cellValue;
                                break;
                            case "D":
                                rowData.Position = cellValue;
                                break;
                            case "E":
                                rowData.Numbers = ParsePhoneNumbers(cellValue);
                                break;
                            case "F":
                                rowData.Numbers.AddRange(ParsePhoneNumbers(cellValue));
                                break;
                            case "G":
                                rowData.Numbers.AddRange(ParsePhoneNumbers(cellValue));
                                break;
                            case "H":
                                rowData.Emails = ParseEmails(cellValue);
                                break;
                            case "I":
                                rowData.Emails.AddRange(ParseEmails(cellValue));
                                break;
                            case "J":
                                rowData.Link = cellValue;
                                break;
                        }
                    }
                    rows.Add(rowData);
                    //Console.WriteLine(rowData.ToString());
                }
            }
            return rows;
        }

        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                SharedStringTablePart stringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (stringTablePart != null)
                {
                    SharedStringItem sharedStringItem = stringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cell.CellValue.Text));
                    return sharedStringItem.InnerText;
                }
            }

            return cell.CellValue?.Text ?? string.Empty;
        }

        private string GetCellColumnName(string cellReference)
        {
            // Remove any digits from the cell reference to get the column name
            string columnName = Regex.Replace(cellReference, @"\d", "");

            return columnName;
        }

        private List<PhoneNumber> ParsePhoneNumbers(string phoneNumbersString)
        {
            List<PhoneNumber> phoneNumbers = new List<PhoneNumber>();

            // Split the phone numbers by a delimiter (e.g., comma or semicolon)
            string[] numbers = phoneNumbersString.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string number in numbers)
            {
                phoneNumbers.Add(new PhoneNumber { Number = number.Trim() });
            }

            return phoneNumbers;
        }

        private List<_Email> ParseEmails(string emailsString)
        {
            List<_Email> emails = new List<_Email>();

            // Split the email addresses by a delimiter (e.g., comma or semicolon)
            string[] emailAddresses = emailsString.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string emailAddress in emailAddresses)
            {
                emails.Add(new _Email { EmailAddress = emailAddress.Trim() });
            }

            return emails;
        }


    }
}