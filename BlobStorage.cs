using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace SchedulingApp.CommonClass
{
    class BlobStorage
    {

        private static string connStr = ConfigurationManager.AppSettings["BlobConnectionString"];
        public string NAsyncUploadFileOnBlob(Stream fileStream, string container, string filePath)
        {
            return UploadFileOnBlob(fileStream, container, filePath).Result;
        }
        public async Task<string> UploadFileOnBlob(Stream fileStream, string container, string filePath)
        {
            try
            {

                // Check whether the connection string can be parsed.
                if (CloudStorageAccount.TryParse(connStr, out CloudStorageAccount storageAccount))
                {
                    // Create the CloudBlobClient that represents the 
                    // Blob storage endpoint for the storage account.
                    CloudBlobClient cloudBlobClient = storageAccount.CreateCloudBlobClient();

                    // Create a container called 'quickstartblobs' and 
                    // append a GUID value to it to make the name unique.
                    CloudBlobContainer cloudBlobContainer =
                        cloudBlobClient.GetContainerReference(container);
                    cloudBlobContainer.CreateIfNotExists();

                    // Set the permissions so the blobs are public.
                    BlobContainerPermissions permissions = new BlobContainerPermissions
                    {
                        PublicAccess = BlobContainerPublicAccessType.Blob
                    };
                    cloudBlobContainer.SetPermissions(permissions);

                    CloudBlockBlob cloudBlockBlob = cloudBlobContainer.GetBlockBlobReference(filePath);
                    cloudBlockBlob.UploadFromStream(fileStream);
                    return cloudBlockBlob.Uri.AbsoluteUri;

                }
                else
                {
                    return "connection not valid";
                }
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in BlobStorage/UploadFileOnBlob function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
                return ex.Message;
            }
        }
        public async Task<string> CopyBlobFile(string fromRoot, string container, string toRoot)
        {
            try
            {

                // Check whether the connection string can be parsed.
                if (CloudStorageAccount.TryParse(connStr, out CloudStorageAccount storageAccount))
                {
                    CloudBlobClient cloudBlobClient = storageAccount.CreateCloudBlobClient();
                    CloudBlobContainer sourceContainer = cloudBlobClient.GetContainerReference(container);
                    CloudBlobContainer targetContainer = cloudBlobClient.GetContainerReference(container);
                    CloudBlockBlob sourceBlob = sourceContainer.GetBlockBlobReference(fromRoot);
                    CloudBlockBlob targetBlob = targetContainer.GetBlockBlobReference(toRoot);
                    await targetBlob.StartCopyAsync(sourceBlob);
                    return targetBlob.Uri.ToString();
                }
                else
                {
                    return "connection not valid";
                }
            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in BlobStorage/CopyBlobFile function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
                return ex.Message;
            }
        }


        public async Task DeleteFileOnBlob(string container, string filePath)
        {
            try
            {

                // Check whether the connection string can be parsed.
                if (CloudStorageAccount.TryParse(connStr, out CloudStorageAccount storageAccount))
                {

                    CloudBlobClient cloudBlobClient = storageAccount.CreateCloudBlobClient();

                    CloudBlobContainer cloudBlobContainer =
                        cloudBlobClient.GetContainerReference(container);

                    // get block blob refarence    
                    CloudBlockBlob blockBlob = cloudBlobContainer.GetBlockBlobReference(filePath);

                    // delete blob from container        
                    await blockBlob.DeleteIfExistsAsync();
                }

            }
            catch (Exception ex)
            {
                CustomLogger.LogError.GetLogger.Log(ex, "Error in BlobStorage/DeleteFileOnBlob function");
                CustomLogger.LogError.GetLogger.Log(ex, ex.Message.ToString());
            }

        }
    }
}