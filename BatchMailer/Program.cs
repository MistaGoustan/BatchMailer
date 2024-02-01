using System;

namespace BatchMailer
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("~~ Batch Mailer ~~");
            Console.WriteLine();
            Console.Write("Press enter to start processing...");
            Console.ReadLine();
            Console.WriteLine();

            var excelReader = new ExcelReader();
            var contacts = excelReader.GetContacts();

            var mailer = new GmailService();

            foreach (var contact in contacts)
            {
                mailer.Send(contact.EmailAddress, contact.Name);
            }

            Console.WriteLine();
            Console.Write("Press enter to close...");
            Console.ReadLine();
        }
    }
}
