using Microsoft.Office.Interop.Outlook;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOtwieranieTest
{
    internal class OutlookEmailSender
    {
        public void CreateEmail(string[][] tableData)
        {
            //Interop.Microsoft.Office.Interop.Outlook z nuget

            Application outlook = new Application();

            // Create a new MailItem object
            MailItem email = (MailItem)outlook.CreateItem(OlItemType.olMailItem);

            // Set the recipients, subject, and body of the email message
            email.To = "recipient@example.com";
            email.Subject = "Table from Jagged Array Example";
            email.BodyFormat = OlBodyFormat.olFormatHTML;

            // Generate a table from the jagged array data
            string tableHtml = "<table>";
            for (int i = 0; i < tableData.Length; i++)
            {
                tableHtml += "<tr>";
                for (int j = 0; j < tableData[i].Length; j++)
                {
                    tableHtml += $"<td>{tableData[i][j]}</td>";
                }
                tableHtml += "</tr>";
            }
            tableHtml += "</table>";

            // Insert the table into the email message body
            email.HTMLBody = tableHtml;

            // Display the email message
            email.Display(true);
        }
    }
}