

using Azure.Identity;
using Microsoft.Graph;

namespace MyApp 
{
    internal class Program
    {
        private static GraphServiceClient? graphClient;
        private static string groupId = "groupid";


        static async Task Main(string[] args)
        {
            var clientId = "clientid";
            var tenantId = "tentantid";

            InteractiveBrowserCredential credential = new(new InteractiveBrowserCredentialOptions()
            {
                ClientId = clientId,
                TenantId = tenantId,
            });
            graphClient = new GraphServiceClient(credential);
            var driveItem = await UploadDocument("documents/Vaardigheden.docx", "General");
            var uploadedPdf = await ConvertDocumentToPdf(driveItem);
            var uploadPdfResult = await UploadDocumentAsPdf(uploadedPdf, $"General/documents/result.pdf");
            if (uploadPdfResult != null)
            {
                Console.WriteLine($"Uploaded PDF file {driveItem?.Name} to {uploadPdfResult.WebUrl}.");
            }
        }




        private async static Task<DriveItem> UploadDocument(string fileName, string sharePointFolder)
        {
            FileStream fileStream = new FileStream(fileName, FileMode.Open);
            return await graphClient.Groups[groupId].Drive.Root.ItemWithPath($"{sharePointFolder}/{fileName}").Content.Request().PutAsync<DriveItem>(fileStream);
        }

        private async static Task<Stream?> ConvertDocumentToPdf(DriveItem item)
        {
            List<QueryOption> options = new List<QueryOption>
            {
                 new QueryOption("$format", "pdf")
            };

            return await graphClient.Groups[groupId].Drive.Items[item.Id].Content.Request(options).GetAsync();

        }

        private async static Task<DriveItem> UploadDocumentAsPdf(Stream? stream, string fileName)
        {
            return await graphClient.Groups[groupId].Drive.Root.ItemWithPath($"{fileName}").Content.Request().PutAsync<DriveItem>(stream);
        }
    }
}