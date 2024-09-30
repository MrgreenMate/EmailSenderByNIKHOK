using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics.Eventing.Reader;

namespace EmailSender
{
    internal class Program
    {
        static void Main()
        {


            // Email template
            string emailTemplate = @"

            Kedves {0},
            !

            Zöldi Máté megbízásából, mellékletben csatolom az e havi elszámolólapod.

            Bármi kérdés esetén keress minket Bizalommal!

            Üdvözlettel:
            A kurva anyád 🙂
             ";

            // List of recipients and their attachments

            List <Recipents> recipients = Recipents.Read();
      

            // SMTP client setup for Outlook
            SmtpClient smtpClient = new SmtpClient("smtp.office365.com")
            {
                Port = 587,
                Credentials = new NetworkCredential("zoldi.mate@uni-obuda.hu", "Zm020327"),
                EnableSsl = true,
            };



            foreach (var recipient in recipients)
            {
                Console.WriteLine(recipient.Name +"      "+recipient.Email+ "      " + recipient.Attachment+"\n");
            }



            Console.WriteLine("Elenőrizd le a listát! Biztos elküldöd az emailt? (I/N)");
            if (Console.ReadLine()== "I")
            {
                foreach (var recipient in recipients)
                {
                    string emailBody = string.Format(emailTemplate, recipient.Name);

                    MailMessage mail = new MailMessage
                    {
                        From = new MailAddress("zoldi.mate@uni-obuda.hu"),
                        Subject = "Teszt email",
                        Body = emailBody,
                        IsBodyHtml = true, // Set to true if your email body is HTML
                    };

                    mail.To.Add(recipient.Email);

                    // Attach the unique file
                    Attachment attachment = new Attachment(recipient.Attachment);
                    mail.Attachments.Add(attachment);

                    try
                    {
                        smtpClient.Send(mail);
                        Console.WriteLine($"Email sent to {recipient.Name} at {recipient.Email}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to send email to {recipient.Name} at {recipient.Email}: {ex.Message}");
                    }
                }

                Console.WriteLine("Emails processing complete.");
            }
            else
            {
                Console.WriteLine("Az emailt nem küldtük el sok puszi!");
            }
        }
       
            

       
    }
}
