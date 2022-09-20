// See https://aka.ms/new-console-template for more information

using BitMiracle.Docotic.Pdf;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.Net.Mail;
using System.Security.Cryptography.X509Certificates;

const string CERT_PATH = ""; //TODO Set path to certificate
const string SMTP_HOST = ""; //TODO Set smtp host
const string LICENSE = ""; //TODO set license key
const string OWNER = ""; //TODO set license owner
const string EMAIL = ""; //TODO set email

if (string.IsNullOrEmpty(EMAIL)) throw new Exception("Set Email");
if (string.IsNullOrEmpty(CERT_PATH)) throw new Exception("Set Path to public key certificate");
if (string.IsNullOrEmpty(SMTP_HOST)) throw new Exception("Set SMTP Host");
if (string.IsNullOrEmpty(LICENSE)) throw new Exception("Set Docotic PDF license key");
if (string.IsNullOrEmpty(OWNER)) throw new Exception("Set Docotic PDF license Owner");

BitMiracle.Docotic.LicenseManager.AddLicenseData(LICENSE, OWNER);

var bytes = Document.Create(container =>
{
    container.Page(page =>
    {
        page.Size(PageSizes.Letter);
        page.Margin(2, Unit.Centimetre);
        page.PageColor(Colors.White);
        page.DefaultTextStyle(x => x.FontSize(20));

        page.Header()
            .Text("Hello PDF!")
            .SemiBold().FontSize(36).FontColor(Colors.Blue.Medium);

        page.Content()
            .PaddingVertical(1, Unit.Centimetre)
            .Column(x =>
            {
                x.Spacing(20);

                x.Item().Text(Placeholders.LoremIpsum());
                x.Item().Image(Placeholders.Image(200, 100));
            });

        page.Footer()
            .AlignCenter()
            .Text(x =>
            {
                x.Span("Page ");
                x.CurrentPageNumber();
            });
    });
}).GeneratePdf();

//Uncomment to save file to disk
//var fileName = $"Encrypted_{DateTime.Now:yyyyMMdd_hhmmss}.pdf";
//var path = Directory.GetCurrentDirectory();
//var filePath = Path.Combine(path, fileName);
byte[] securePdf;
int securePdfLength;
using (var pdf = new PdfDocument(bytes))
{
    var cert = new X509Certificate2(CERT_PATH);
    var handler = new PdfPublicKeyEncryptionHandler(cert);
    var permissions = new PdfPermissions
    {
        Flags = PdfPermissionFlags.AssembleDocument | PdfPermissionFlags.PrintDocument | PdfPermissionFlags.PrintFaithfulCopy | PdfPermissionFlags.ModifyContents
        | PdfPermissionFlags.CopyContents | PdfPermissionFlags.ModifyAnnotations | PdfPermissionFlags.FillFormFields | PdfPermissionFlags.ExtractContents
    };

    //handler.AddRecipient(cert, permissions);

    var saveOptions = new PdfSaveOptions { EncryptionHandler = handler };

    using (var ms = new MemoryStream())
    {
        pdf.Save(ms, saveOptions);
        securePdfLength = (int)ms.Length;
        securePdf = ms.ToArray();

        //Uncomment to save file to disk
        //pdf.Save(filePath, saveOptions);
    }
}

var mailaddress = new MailAddress(EMAIL);
var mailMessage = new MailMessage(mailaddress, mailaddress)
{
    Subject = "Encrypted PDF"
};

//Uncomment to attach file from disk
//var fileBytes = File.ReadAllBytes(filePath);
//var attachment = new Attachment(new MemoryStream(fileBytes, 0, fileBytes.Length), fileName);

var attachment = new Attachment(new MemoryStream(securePdf, 0, securePdfLength), "file.pdf");
mailMessage.Attachments.Add(attachment);
using var smtp = new SmtpClient(SMTP_HOST);
smtp.Send(mailMessage);
