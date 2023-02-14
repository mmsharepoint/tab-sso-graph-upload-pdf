﻿using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.Graph;
using Microsoft.Identity.Web;
using System.IO;
using System.Net;
using TabSSOGraphUploadPDF.Models;

namespace TabSSOGraphUploadPDF.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UploadController : ControllerBase
    {
        private readonly GraphServiceClient _graphClient;
        private readonly ITokenAcquisition _tokenAcquisition;
        private readonly ILogger<UploadController> _logger;
        public UploadController(ITokenAcquisition tokenAcquisition, GraphServiceClient graphClient, ILogger<UploadController> logger)
        {
            _tokenAcquisition = tokenAcquisition;
            _graphClient = graphClient;
            _logger = logger;
        }
        // api/<controller>/GetMimeMessage
        [HttpPost]
        [DisableRequestSizeLimit] //<======= add this line
        [RequestFormLimits(ValueLengthLimit = int.MaxValue, MultipartBodyLengthLimit = int.MaxValue)]
        public async Task<ActionResult<string>> Post([FromForm] UploadRequest fileUpload)
        {
            string accessToken = await GetAccessToken();

            string fileName = fileUpload.Name;
            string siteUrl = fileUpload.SiteUrl;
            _logger.LogInformation($"Received file {fileUpload.file.FileName} with size in bytes {fileUpload.file.Length}");
            string userID = User.GetObjectId(); //   Claims["preferred_username"];
            DriveItem uploadResult = await this._graphClient.Users[userID]
                                                    .Drive.Root
                                                    .ItemWithPath(fileUpload.file.FileName)
                                                    .Content.Request()
                                                    .PutAsync<DriveItem>(fileUpload.file.OpenReadStream());

            Stream pdfFile = await GetPDF(userID, uploadResult.Id);
            string pdfFileUrl = await UploadPDF(userID, fileUpload.file.FileName, pdfFile);
            DeleteTempFile(userID, uploadResult.Id);
            return Ok(pdfFileUrl);
            //return Ok(uploadResult.WebUrl);
        }

        private async Task<string> GetAccessToken()
        {
            _logger.LogInformation($"Authenticated user: {User.GetDisplayName()}");

            try
            {
                // TEMPORARY
                // Get a Graph token via OBO flow
                var token = await _tokenAcquisition
                    .GetAccessTokenForUserAsync(new[]{
                        "Files.ReadWrite", "Sites.ReadWrite.All" });

                // Log the token
                _logger.LogInformation($"Access token for Graph: {token}");
                return token;
            }
            catch (MicrosoftIdentityWebChallengeUserException ex)
            {
                _logger.LogError(ex, "Consent required");
                // This exception indicates consent is required.
                // Return a 403 with "consent_required" in the body
                // to signal to the tab it needs to prompt for consent
                return "";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error occurred");
                return "";
            }
        }

        private async Task<Stream> GetPDF(string userID, string itemID)
        {
            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("format", "PDF")
            };
            Stream pdfResult = await this._graphClient.Users[userID]
                                                    .Drive.Items[itemID]
                                                    .Content
                                                    .Request(queryOptions)
                                                    .GetAsync();
            return pdfResult;
        }

        private async Task<string> UploadPDF(string userID, string orgFileName, Stream fileStream)
        {
            string pdfFileName = Path.GetFileNameWithoutExtension(orgFileName);
            pdfFileName += ".pdf";
            DriveItem uploadResult = await this._graphClient.Users[userID]
                                                    .Drive.Root
                                                    .ItemWithPath(pdfFileName)
                                                    .Content.Request()
                                                    .PutAsync<DriveItem>(fileStream);
            return uploadResult.WebUrl;
        }

        private async Task DeleteTempFile(string userID, string itemID)
        {
            await this._graphClient.Users[userID]
                        .Drive.Items[itemID]
                        .Request()
                        .DeleteAsync();
        }
    }
}
