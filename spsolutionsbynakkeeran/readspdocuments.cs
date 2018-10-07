using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Microsoft.Office.Interop.Word;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Xml;

namespace spsolutionsbynakkeeran
{
    public static class readspdocuments
    {
        [FunctionName("readspdocuments")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            // Credentials for getting authenticated context using auth manager of OfficeDevPnP Dll
            // Since its a POC, I have used direct creds.
            // You could use other authentication modes for getting the context for PRODUCTION ENV usage.
	        string siteUrl = "https://nakkeerann.sharepoint.com/sites/teamsite";
	        string userName = "nav@nakkeerann.onmicrosoft.com";
	        string password = "*****";
	        OfficeDevPnP.Core.AuthenticationManager authManager = new OfficeDevPnP.Core.AuthenticationManager();
	        // parse query parameter
	        string filePath = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "filePath", true) == 0)
                .Value;
            
	        try
	        {
                // context using auth manager
		        using (var clientContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, userName, password))
		        {
			        Web web = clientContext.Site.RootWeb;
                    Microsoft.SharePoint.Client.File file = web.GetFileByUrl(filePath);
                    
                    var data = file.OpenBinaryStream();
                    string content = null;
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        clientContext.Load(file);
                        clientContext.ExecuteQuery();

                        if (data != null && data.Value != null)
                        {
                            data.Value.CopyTo(memoryStream);
                            memoryStream.Seek(0, SeekOrigin.Begin);
                            
                            // Function extracts the document content
                            content = ExtractContentFromWordDocument(memoryStream);
                            
                            // Function responds back with extracted document content
                            return req.CreateResponse(HttpStatusCode.OK, content);
                        }
                    }
                    // Function responds back
                    return req.CreateResponse(HttpStatusCode.BadRequest, "Unable to process file or no content present");
                }
            }
	        catch (Exception ex)
	        {
		        log.Info("Error Message: " + ex.Message);
                return req.CreateResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        public static string ExtractContentFromWordDocument(MemoryStream filePath)
        {
            // open xml namespace format for processing documents
            string xmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            StringBuilder textBuilder = new StringBuilder();
            using (WordprocessingDocument processDocument = WordprocessingDocument.Open(filePath, false))
            {
                NameTable nameTable = new NameTable();
                XmlNamespaceManager xmlNamespaceManager = new XmlNamespaceManager(nameTable);
                xmlNamespaceManager.AddNamespace("w", xmlNamespace);

                // Extract all paragraphs from document XML
                XmlDocument xmlDocument = new XmlDocument(nameTable);
                xmlDocument.Load(processDocument.MainDocumentPart.GetStream());
                XmlNodeList paragraphNodes = xmlDocument.SelectNodes("//w:p", xmlNamespaceManager);

                // Parse through each paragraph nodes
                foreach (XmlNode paragraphNode in paragraphNodes)
                {
                    // Get only text nodes, excluding the other formatting options
                    XmlNodeList textNodes = paragraphNode.SelectNodes(".//w:t", xmlNamespaceManager);

                    // Append only text content to the custom string builder
                    foreach (XmlNode textNode in textNodes)
                    {
                        textBuilder.Append(textNode.InnerText);
                    }
                    textBuilder.AppendLine();
                }
            }
            return textBuilder.ToString();
        }
    }
}
