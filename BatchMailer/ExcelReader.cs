using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace BatchMailer
{
    internal class ExcelReader
    {
        // Stackoverflow: How to read from XLSX (Excel)?
        // https://stackoverflow.com/questions/33302235/how-to-read-from-xlsx-excel
        public List<Contact> GetContacts()
        {
            var contacts = new List<Contact>();

            //string path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + @"\Contacts.xlsx";
            string path = "Contacts.xlsx";
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(path)))
            {
                var myWorksheet = xlPackage.Workbook.Worksheets.First();
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;

                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                for (int row = 2; row <= rowCount; row++) // Start at row 2 to skip the column headers
                {
                    var newContact = new Contact()
                    {
                        Name = worksheet.Cells[row, 1].Value?.ToString().Trim(),
                        EmailAddress = worksheet.Cells[row, 2].Value?.ToString().Trim(),
                    };

                    if (!IsValidContact(newContact, row))
                    {
                        continue;
                    }

                    contacts.Add(newContact);

                    Console.WriteLine($"[SUCCESS] Name: {newContact.Name} - Email: {newContact.EmailAddress}");
                }
            }

            return contacts;
        }

        private bool HasValidName(string name, int row)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            return true;
        }

        private bool HasValidEmail(string email, int row)
        {
            if (string.IsNullOrWhiteSpace(email))
            {
                return false;
            }

            Regex regex = new Regex(@"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,3})+)$");
            Match match = regex.Match(email);

            if (!match.Success)
            {
                Console.WriteLine($"[DATA ERROR] Row #{row} has an invalid email address.");
            }

            return match.Success;
        }

        private bool IsValidContact(Contact contact, int row)
        {
            var hasName = HasValidName(contact.Name, row);
            var hasEmail = HasValidEmail(contact.EmailAddress, row);

            if (hasName && !hasEmail)
            {
                Console.WriteLine($"[DATA ERROR] Row #{row} is missing, or has an invalid email address.");

                return false;
            }
            if (!hasName && hasEmail)
            {
                Console.WriteLine($"[DATA ERROR] Row #{row} is missing display name.");

                return false;
            }
            if (!hasName && !hasEmail)
            {
                Console.WriteLine($"[BLANK] Row #{row}");

                return false;
            }

            return true;
        }
    }
}