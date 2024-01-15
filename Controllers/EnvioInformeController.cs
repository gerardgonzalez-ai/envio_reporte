using Microsoft.AspNetCore.Mvc;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Reflection.Metadata;
using Aspose.Pdf.Text;
using Aspose.Pdf;
using Aspose.Pdf.Operators;
using Azure.Core;

namespace EnvioInforme.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class EnvioInformeController : ControllerBase
    {

        private readonly IConfiguration _config;
        private readonly ILogger<EnvioInformeController> _logger;
        private IWebHostEnvironment _environment;


        private string? tenantId;
        private string? clientId;
        private string? clientSecret;
        private string? fromAddress;
        private string? toAddress;
        private string? siteID;
        private string? driveID;
        private string? destinationLibraryPath;

        public EnvioInformeController(IConfiguration config, ILogger<EnvioInformeController> logger, IWebHostEnvironment environment)
        {
            _logger = logger;
            _config = config;
            tenantId = _config.GetValue<string>("tenantId");
            clientId = _config.GetValue<string>("clientId");
            clientSecret = _config.GetValue<string>("clientSecret");
            fromAddress = _config.GetValue<string>("fromAddress");
            toAddress = _config.GetValue<string>("toAddress");
            siteID = _config.GetValue<string>("siteID");
            driveID = _config.GetValue<string>("driveID");
            destinationLibraryPath = _config.GetValue<string>("destinationLibraryPath");
            _environment = environment;
        }

        [HttpPost(Name = "EnviarInforme")]
        public async Task<OperationResponse> EnviarInforme(OperationRequest request)
        {
            OperationResponse result = new OperationResponse();
            result.OperationResult = false;

            //TODO: 1.- Recuperar Imagenes
            //TODO: 2.- Armar el PDF
            //TODO: 3.- Obtener el PDF en B64

            string reportB64 = GeneratePDF(request); 
    
            if (request != null && request.UserName != null && request.ReportName != null && reportB64 != null)
            {
                DateTime utcNow = DateTime.UtcNow;
                TimeZoneInfo chileTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific SA Standard Time");
                DateTime fechaProceso = TimeZoneInfo.ConvertTimeFromUtc(utcNow, chileTimeZone);

                request.ReportName = request.ReportName.Replace(".pdf", " ") + fechaProceso.ToString("dd-MM-yyy HHmmss");
                request.ReportName += ".pdf";

                String urlFile = await uploadFileSPO(request, reportB64);

                if (urlFile != null && urlFile != "")
                {
                    string subject = "App Informes >> Notificación nuevo informe " + fechaProceso.ToString("dd-MM-yyy HH:mm:ss");
                    string content = "<body><div><span style=\"font-family:Arial, Helvetica, sans-serif;\"><b>Se ha agregado un nuevo Informe:</b></span><br><br><table border=1> <tbody><tr style=\"background-color:#003a70;\"><td style=\"width: 400px; height: 21px; text-align: center; color:#FFF\" colspan=1><b>Datos del Informe<b><td> </tr><tr><td><b>Fecha:</b></td><td style=\"width: 250px; height: 21px;\">@FechaInforme</td></tr><tr><td><b>Autor:</b></td><td style=\"width: 250px; height: 21px;\">@NombreAutor</td></tr><tr><td><b>URL Informe</b></td><td style=\"width: 250px; height: 21px;\"><a href=\"@UrlInforme\">@NombreInforme</a></td></tr></tbody></table></div></body>";
                    content = content.Replace("@FechaInforme", fechaProceso.ToString("dd-MM-yyy HH:mm:ss"));
                    content = content.Replace("@NombreAutor", request.UserName);
                    content = content.Replace("@UrlInforme", urlFile);
                    content = content.Replace("@NombreInforme", request.ReportName);

                    //result.OperationResult = await SendEmailCBM(subject, content, request.UserName, urlFile);
                }

            }
            return result;
        }

        private async Task<String> uploadFileSPO(OperationRequest request, String reportB64)
        {
            String urlFile = "";

            try
            {
                var confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
                .Build();

                var authResult = await confidentialClientApplication.AcquireTokenForClient(new[] { "https://graph.microsoft.com/.default" })
                    .ExecuteAsync();

                var graphServiceClient = new GraphServiceClient(new DelegateAuthenticationProvider(requestMessage =>
                {
                    requestMessage.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                    return Task.CompletedTask;
                }));

                if (reportB64 != null)
                {
                    // Read the file
                    byte[] fileBytes = Convert.FromBase64String(reportB64);

                    // Get the destination library drive item
                    var driveItem = await graphServiceClient
                   .Sites[siteID]
                   .Drives[driveID]
                   .Root
                   .ItemWithPath(destinationLibraryPath + request.ReportName)
                   .Content
                   .Request()
                   .PutAsync<DriveItem>(new MemoryStream(fileBytes));

                    urlFile = driveItem.WebUrl;

                }

            }
            catch (Exception e)
            {
                _logger.LogError("Subiendo archivo SharePoint: " + e.Message);
            }

            return urlFile;
        }

        private async Task<bool> SendEmailCBM(string subject, string content, string userName, string urlInforme)
        {
            bool flagSendEmail = false;

            ClientSecretCredential credential = new(tenantId, clientId, clientSecret);
            GraphServiceClient graphClient = new(credential);

            Message message = new()
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = content
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = toAddress
                        }
                    }
                }
            };

            bool saveToSentItems = true;

            try
            {
                await graphClient.Users[fromAddress]
               .SendMail(message, saveToSentItems)
               .Request()
               .PostAsync();

                flagSendEmail = true;
            }
            catch (Exception e)
            {
                _logger.LogError("Envío mail: " + e.Message);
            }


            return flagSendEmail;
        }

        private string GeneratePDF(OperationRequest request)
        {
            string base64String = "";

            string rootPath = _environment.WebRootPath;
            string folderPath = Path.Combine(rootPath, "Docs");
            
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(Path.Combine(folderPath, $"{request.ReportType}.pdf"));

            if (request.ReportType != null && request.ReportType.ToLower() == "rfc")
            {
                ReplaceTextInPDF(pdfDocument, "AVAL", request.UserName);
                ReplaceTextInPDF(pdfDocument, "BVAL", request.RfcDescription);
                ReplaceTextInPDF(pdfDocument, "CVAL", request.RfcObservation);

                int numberOfTables = request.TotalImage % 4;
                for (int i = 1; i <= numberOfTables; i++)
                {

                    int index;
                    if (i == 1)
                         index = i;
                    else
                        index = i+3;

                    Aspose.Pdf.Table newTable = new Aspose.Pdf.Table();
                    newTable.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, .1f, Aspose.Pdf.Color.FromRgb(System.Drawing.Color.White));
                    newTable.DefaultCellBorder = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, .1f, Aspose.Pdf.Color.FromRgb(System.Drawing.Color.White));
                    newTable.ColumnWidths = "120 120 120 120";
                    newTable.DefaultCellPadding = new MarginInfo(2, 2, 2, 2);
                    newTable.Margin = new MarginInfo(2, 2, 2, 2);


                    Aspose.Pdf.Row row1 = newTable.Rows.Add();

                    if (System.IO.File.Exists(Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index}.png")))
                    {
                        Aspose.Pdf.Image img1 = new Aspose.Pdf.Image();
                        img1.File = Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index}.png");
                        img1.FixWidth = 100;
                        img1.FixHeight = 50;

                        Aspose.Pdf.Cell cell1 = row1.Cells.Add();
                        cell1.Paragraphs.Add(img1);
                    }
                    else
                    {
                        Aspose.Pdf.Cell cell1 = row1.Cells.Add("No Image Found");
                    }

                    if (System.IO.File.Exists(Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 1}.png")))
                    {
                        Aspose.Pdf.Image img2 = new Aspose.Pdf.Image();
                        img2.File = Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 1}.png");
                        img2.FixWidth = 100;
                        img2.FixHeight = 50;

                        Aspose.Pdf.Cell cell2 = row1.Cells.Add();
                        cell2.Paragraphs.Add(img2);
                    }
                    else
                    {
                        Aspose.Pdf.Cell cell1 = row1.Cells.Add("No Image Found");
                    }

                    if (System.IO.File.Exists(Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 2}.png")))
                    {
                        Aspose.Pdf.Image img3 = new Aspose.Pdf.Image();
                        img3.File = Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 2}.png");
                        img3.FixWidth = 100;
                        img3.FixHeight = 50;

                        Aspose.Pdf.Cell cell3 = row1.Cells.Add();
                        cell3.Paragraphs.Add(img3);
                    }
                    else
                    {
                        Aspose.Pdf.Cell cell1 = row1.Cells.Add("No Image Found");
                    }

                    if (System.IO.File.Exists(Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 3}.png")))
                    {
                        Aspose.Pdf.Image img4 = new Aspose.Pdf.Image();
                        img4.File = Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 3}.png");
                        img4.FixWidth = 100;
                        img4.FixHeight = 50;

                        Aspose.Pdf.Cell cell4 = row1.Cells.Add();
                        cell4.Paragraphs.Add(img4);
                    }
                    else
                    {
                        Aspose.Pdf.Cell cell1 = row1.Cells.Add("No Image Found");
                    }


                    Page lastPage = pdfDocument.Pages[pdfDocument.Pages.Count];

                    //double newYCoordinate = lastPage.Rect.Height - newTable.GetHeight();

                    if (i == 1)
                        newTable.Margin = new MarginInfo { Top = 100 };
                    else
                        newTable.Margin = new MarginInfo { Top = 20 };

                    lastPage.Paragraphs.Add(newTable);


                }

                string saveFilePath = Path.Combine(rootPath, "Temp", $"Output_{request.ReportName}_{Guid.NewGuid().ToString()}.pdf");
                pdfDocument.Save(saveFilePath);
                byte[] pdfBytes = System.IO.File.ReadAllBytes(saveFilePath);
                base64String = Convert.ToBase64String(pdfBytes);


            }

            if (request.ReportType != null && request.ReportType.ToLower() == "rfdc")
            {
                ReplaceTextInPDF(pdfDocument, "AVAL", request.UserName);
                ReplaceTextInPDF(pdfDocument, "BVAL", request.RfdcOriginRefrigerator);
                ReplaceTextInPDF(pdfDocument, "CVAL", request.RfdcDispatchType);
                ReplaceTextInPDF(pdfDocument, "DVAL", request.RfdcCustomer);
                ReplaceTextInPDF(pdfDocument, "EVAL", request.RfdcDestination);
                ReplaceTextInPDF(pdfDocument, "FVAL", request.RfdcSpecie);
                ReplaceTextInPDF(pdfDocument, "GVAL", request.RfdcStatus);
                ReplaceTextInPDF(pdfDocument, "HVAL", request.RfdcObservation);

                int numberOfTables = request.TotalImage / 4;
                for (int i = 1; i <= numberOfTables; i++)
                {

                    int index;
                    if (i == 1)
                        index = i;
                    else
                        index = i + 3;

                    Aspose.Pdf.Table newTable = new Aspose.Pdf.Table();
                    newTable.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, .1f, Aspose.Pdf.Color.FromRgb(System.Drawing.Color.White));
                    newTable.DefaultCellBorder = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, .1f, Aspose.Pdf.Color.FromRgb(System.Drawing.Color.White));
                    newTable.ColumnWidths = "120 120 120 120";
                    newTable.DefaultCellPadding = new MarginInfo(2, 2, 2, 2);
                    newTable.Margin = new MarginInfo(2, 2, 2, 2);

                    Aspose.Pdf.Row row1 = newTable.Rows.Add();

                    if (System.IO.File.Exists(Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index}.png")))
                    {
                        Aspose.Pdf.Image img1 = new Aspose.Pdf.Image();
                        img1.File = Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index}.png");
                        img1.FixWidth = 100;
                        img1.FixHeight = 50;

                        Aspose.Pdf.Cell cell1 = row1.Cells.Add();
                        cell1.Paragraphs.Add(img1);
                    }
                    else
                    {
                        Aspose.Pdf.Cell cell1 = row1.Cells.Add("No Image Found");
                    }

                    if (System.IO.File.Exists(Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 1}.png")))
                    {
                        Aspose.Pdf.Image img2 = new Aspose.Pdf.Image();
                        img2.File = Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 1}.png");
                        img2.FixWidth = 100;
                        img2.FixHeight = 50;

                        Aspose.Pdf.Cell cell2 = row1.Cells.Add();
                        cell2.Paragraphs.Add(img2);
                    }
                    else
                    {
                        Aspose.Pdf.Cell cell1 = row1.Cells.Add("No Image Found");
                    }

                    if (System.IO.File.Exists(Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 2}.png")))
                    {
                        Aspose.Pdf.Image img3 = new Aspose.Pdf.Image();
                        img3.File = Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 2}.png");
                        img3.FixWidth = 100;
                        img3.FixHeight = 50;

                        Aspose.Pdf.Cell cell3 = row1.Cells.Add();
                        cell3.Paragraphs.Add(img3);
                    }
                    else
                    {
                        Aspose.Pdf.Cell cell1 = row1.Cells.Add("No Image Found");
                    }

                    if (System.IO.File.Exists(Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 3}.png")))
                    {
                        Aspose.Pdf.Image img4 = new Aspose.Pdf.Image();
                        img4.File = Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 3}.png");
                        img4.FixWidth = 100;
                        img4.FixHeight = 50;

                        Aspose.Pdf.Cell cell4 = row1.Cells.Add();
                        cell4.Paragraphs.Add(img4);
                    }
                    else
                    {
                        Aspose.Pdf.Cell cell1 = row1.Cells.Add("No Image Found");
                    }


                    Page lastPage = pdfDocument.Pages[pdfDocument.Pages.Count];

                    //double newYCoordinate = lastPage.Rect.Height - newTable.GetHeight();

                    if (i == 1)
                        newTable.Margin = new MarginInfo { Top = 270 };
                    else
                        newTable.Margin = new MarginInfo { Top = 20 };

                    lastPage.Paragraphs.Add(newTable);


                }

                string saveFilePath = Path.Combine(rootPath, "Temp", $"Output_{request.ReportName}_{Guid.NewGuid().ToString()}.pdf");
                pdfDocument.Save(saveFilePath);
                byte[] pdfBytes = System.IO.File.ReadAllBytes(saveFilePath);
                base64String = Convert.ToBase64String(pdfBytes);

            }

            if (request.ReportType != null && request.ReportType.ToLower() == "rff")
            {
                ReplaceTextInPDF(pdfDocument, "AVAL", request.UserName);
                ReplaceTextInPDF(pdfDocument, "BVAL", request.RffType);
                ReplaceTextInPDF(pdfDocument, "CVAL", request.RffDispatchDate);
                ReplaceTextInPDF(pdfDocument, "DVAL", request.RffDestination);
                ReplaceTextInPDF(pdfDocument, "EVAL", request.RffSpecie);
                ReplaceTextInPDF(pdfDocument, "FVAL", request.RffPortDeparture);
                ReplaceTextInPDF(pdfDocument, "GVAL", request.RffPortDeparture);
                ReplaceTextInPDF(pdfDocument, "HVAL", request.RffPortDestination);
                ReplaceTextInPDF(pdfDocument, "IVAL", request.RffObservation);

                int numberOfTables = request.TotalImage % 4;
                for (int i = 1; i <= numberOfTables; i++)
                {

                    int index;
                    if (i == 1)
                        index = i;
                    else
                        index = i + 3;

                    Aspose.Pdf.Table newTable = new Aspose.Pdf.Table();
                    newTable.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, .1f, Aspose.Pdf.Color.FromRgb(System.Drawing.Color.White));
                    newTable.DefaultCellBorder = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, .1f, Aspose.Pdf.Color.FromRgb(System.Drawing.Color.White));
                    newTable.ColumnWidths = "120 120 120 120";
                    newTable.DefaultCellPadding = new MarginInfo(2, 2, 2, 2);
                    newTable.Margin = new MarginInfo(2, 2, 2, 2);

                    Aspose.Pdf.Row row1 = newTable.Rows.Add();

                    if (System.IO.File.Exists(Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index}.png")))
                    {
                        Aspose.Pdf.Image img1 = new Aspose.Pdf.Image();
                        img1.File = Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index}.png");
                        img1.FixWidth = 100;
                        img1.FixHeight = 50;

                        Aspose.Pdf.Cell cell1 = row1.Cells.Add();
                        cell1.Paragraphs.Add(img1);
                    }
                    else
                    {
                        Aspose.Pdf.Cell cell1 = row1.Cells.Add("No Image Found");
                    }

                    if (System.IO.File.Exists(Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 1}.png")))
                    {
                        Aspose.Pdf.Image img2 = new Aspose.Pdf.Image();
                        img2.File = Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 1}.png");
                        img2.FixWidth = 100;
                        img2.FixHeight = 50;

                        Aspose.Pdf.Cell cell2 = row1.Cells.Add();
                        cell2.Paragraphs.Add(img2);
                    }
                    else
                    {
                        Aspose.Pdf.Cell cell1 = row1.Cells.Add("No Image Found");
                    }

                    if (System.IO.File.Exists(Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 2}.png")))
                    {
                        Aspose.Pdf.Image img3 = new Aspose.Pdf.Image();
                        img3.File = Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 2}.png");
                        img3.FixWidth = 100;
                        img3.FixHeight = 50;

                        Aspose.Pdf.Cell cell3 = row1.Cells.Add();
                        cell3.Paragraphs.Add(img3);
                    }
                    else
                    {
                        Aspose.Pdf.Cell cell1 = row1.Cells.Add("No Image Found");
                    }

                    if (System.IO.File.Exists(Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 3}.png")))
                    {
                        Aspose.Pdf.Image img4 = new Aspose.Pdf.Image();
                        img4.File = Path.Combine(rootPath, "Images", $"{request.ReportGuid}_{index + 3}.png");
                        img4.FixWidth = 100;
                        img4.FixHeight = 50;

                        Aspose.Pdf.Cell cell4 = row1.Cells.Add();
                        cell4.Paragraphs.Add(img4);
                    }
                    else
                    {
                        Aspose.Pdf.Cell cell1 = row1.Cells.Add("No Image Found");
                    }


                    Page lastPage = pdfDocument.Pages[pdfDocument.Pages.Count];

                    //double newYCoordinate = lastPage.Rect.Height - newTable.GetHeight();

                    if (i == 1)
                        newTable.Margin = new MarginInfo { Top = 270 };
                    else
                        newTable.Margin = new MarginInfo { Top = 20 };

                    lastPage.Paragraphs.Add(newTable);


                }

                string saveFilePath = Path.Combine(rootPath, "Temp", $"Output_{request.ReportName}_{Guid.NewGuid().ToString()}.pdf");
                pdfDocument.Save(saveFilePath);
                byte[] pdfBytes = System.IO.File.ReadAllBytes(saveFilePath);
                base64String = Convert.ToBase64String(pdfBytes);
            }

            return base64String;

        }

        private void ReplaceTextInPDF(Aspose.Pdf.Document pdfDocument, string textToReplace, string valueToReplace)
        {
            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber(textToReplace, new TextSearchOptions(true));

            pdfDocument.Pages.Accept(textFragmentAbsorber);

            TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;

            foreach (TextFragment textFragment in textFragmentCollection)
            {
                textFragment.Text = valueToReplace;
            }
        }

    }
}