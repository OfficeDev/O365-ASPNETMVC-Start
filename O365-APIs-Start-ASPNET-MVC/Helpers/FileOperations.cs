using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office365.SharePoint.FileServices;
using Microsoft.Office365.SharePoint.CoreServices.Extensions;
using model = O365_APIs_Start_ASPNET_MVC.Models;
using System.Threading.Tasks;
using Microsoft.OData.Client;
using Microsoft.OData.Core;
using System.Diagnostics;
using System.Text;
using System.IO;

namespace O365_APIs_Start_ASPNET_MVC.Helpers
{
    public class FileOperations
    {
        /// <summary>
        /// Performs a search of the default Documents folder. 
        /// </summary>
        /// <returns>A collection of information that describes files and folders.</returns>
        internal async Task<List<model.FileObject>> GetMyFilesAsync()
        {
            var sharePointClient = await AuthenticationHelper.EnsureSharePointClientCreatedAsync("MyFiles");
            //IOrderedEnumerable<IFileSystemItem> files = null;

            List<model.FileObject> returnResults = new List<model.FileObject>();

            try
            {
                // Performs a search of the default Documents folder.
                // You could also specify other folders using the syntax: var filesResults = await _client.Files["folder_name"].ExecuteAsync();
                // This results in a call to the service.

                var filesResults = await sharePointClient.Files.ExecuteAsync();

                var files = filesResults.CurrentPage;

                foreach (IItem fileItem in files)
                {
                    // The item to add to the result set.
                    model.FileObject modelFile = new model.FileObject(fileItem);

                    returnResults.Add(modelFile);
                }
            }
            catch (ODataErrorException)
            {
                return null;
            }
            catch (DataServiceQueryException)
            {
                return null;
            }
            catch (MissingMethodException e)
            {
                Debug.Write(e.Message);
            }

            return returnResults;
        }

        //<summary>
        //Creates a new file named demo.txt in the default document library.
        //</summary>
        //<returns>A Boolean value that indicates whether the new text file was successfully created.</returns>
        internal async Task<String> CreateNewTextFileAsync()
        {
            //bool isSuccess = false;
            String newID = string.Empty;
            var sharePointClient = await AuthenticationHelper.EnsureSharePointClientCreatedAsync("MyFiles");

            try
            {
                // First check whether demo.txt already exists. If it exists, delete it.
                // If it doesn't exist, swallow the error.
                IItem item = await sharePointClient.Files.GetByPathAsync("demo.txt");
                await item.DeleteAsync();
            }
            catch (ODataErrorException)
            {
                // fail silently because demo.txt doesn't exist.
            }
           
            try
            {
                
                // In this example, we'll create a simple text file and write the current timestamp into it. 
                string createdTime = "Created at " + DateTime.Now.ToLocalTime().ToString();
                byte[] bytes = Encoding.UTF8.GetBytes(createdTime);

                using (MemoryStream stream = new MemoryStream(bytes))
                {
                    // File is called demo.txt. If it already exists, we'll get an exception. 
                    Microsoft.Office365.SharePoint.FileServices.File newFile = new Microsoft.Office365.SharePoint.FileServices.File
                    {
                        Name = "demo.txt"
                    };

                    // Create the empty file.
                    await sharePointClient.Files.AddItemAsync(newFile);
                    newID = newFile.Id;

                    // Upload the file contents.
                    await sharePointClient.Files.GetById(newFile.Id).ToFile().UploadAsync(stream);
                }
            }

            // ODataErrorException can be thrown when you try to create a file that already exists.
            catch (Microsoft.Data.OData.ODataErrorException)
            {
                //isSuccess = false;
            }

            return newID;
        }

        /// <summary>
        /// Deletes the selected item or folder from the ListBox.
        /// </summary>
        /// <returns>A Boolean value that indicates whether the file or folder was successfully deleted.</returns>
        internal async Task<bool?> DeleteFileOrFolderAsync(string selectedItemID)
        {
            bool? isSuccess = false;

            try
            {
                // Make sure we have a reference to the SharePoint client
                var spClient = await AuthenticationHelper.EnsureSharePointClientCreatedAsync("MyFiles");

                // Get the file to be removed from the SharePoint service. This results in a call to the service.
                IItemFetcher thisItemFetcher = spClient.Files.GetById(selectedItemID);
                IFileFetcher thisFileFetcher = thisItemFetcher.ToFile();
                var thisFile = await thisFileFetcher.ExecuteAsync();

                // Delete the file or folder. This results in a call to the service.
                await thisFile.DeleteAsync();

                isSuccess = true;
            }
            catch (Microsoft.Data.OData.ODataErrorException)
            {
                isSuccess = null;
            }
            catch (NullReferenceException)
            {
                isSuccess = null;
            }

            return isSuccess;
        }

        /// <summary>
        /// Reads the contents of a text file and displays the results in a TextBox.
        /// </summary>
        /// <param name="_selectedFileObject">The file selected in the ListBox.</param>
        /// <returns>A Boolean value that indicates whether the text file was successfully read.</returns>
        internal async Task<object[]> ReadTextFileAsync(string selectedItemID)
        {

            string fileContents = string.Empty;
            object[] results = new object[] { fileContents, false };

            try
            {
                // Make sure we have a reference to the SharePoint client
                var spClient = await AuthenticationHelper.EnsureSharePointClientCreatedAsync("MyFiles");

                // Get a handle on the selected item.
                IItemFetcher thisItemFetcher = spClient.Files.GetById(selectedItemID);
                IFileFetcher thisFileFetcher = thisItemFetcher.ToFile();
                var myFile = await thisFileFetcher.ExecuteAsync();

                // Check that the selected item is a .txt file.
                if (!myFile.Name.EndsWith(".txt") && !myFile.Name.EndsWith(".xml"))
                {
                    results[0] = string.Empty;
                    results[1] = false;
                    return results;
                }

                Microsoft.Office365.SharePoint.FileServices.File file = myFile as Microsoft.Office365.SharePoint.FileServices.File;

                // Download the text file and put the results into a string. This results in a call to the service.
                using (Stream stream = await file.DownloadAsync())
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        results[0] = await reader.ReadToEndAsync();

                        results[1] = true;
                    }
                }
            }
            catch (NullReferenceException)
            {
                results[1] = false;
            }
            catch (ArgumentException)
            {
                results[1] = false;
            }

            return results;
        }

        /// <summary>
        /// Update the currently selected item by appending new text.
        /// </summary>
        /// <param name="_selectedFileObject">The file selected in the ListBox.</param>
        /// <param name="fileText">The updated text contents of the file.</param>
        /// <returns>A Boolean value that indicates whether the text file was successfully updated.</returns>
        internal async Task<bool> UpdateTextFileAsync(string selectedItemID, string fileText)
        {
            Microsoft.Office365.SharePoint.FileServices.File file;
            byte[] byteArray;
            bool isSuccess = false;

            try
            {
                // Make sure we have a reference to the SharePoint client
                var spClient = await AuthenticationHelper.EnsureSharePointClientCreatedAsync("MyFiles");

                // Get a handle on the selected item.
                IItemFetcher thisItemFetcher = spClient.Files.GetById(selectedItemID);
                IFileFetcher thisFileFetcher = thisItemFetcher.ToFile();
                var myFile = await thisFileFetcher.ExecuteAsync();

                file = myFile as Microsoft.Office365.SharePoint.FileServices.File;
                string updateTime = "\n\r\n\rLast update at " + DateTime.Now.ToLocalTime().ToString();
                byteArray = Encoding.UTF8.GetBytes(fileText + updateTime);

                using (MemoryStream stream = new MemoryStream(byteArray))
                {
                    // Update the file. This results in a call to the service.
                    await file.UploadAsync(stream);
                    isSuccess = true; // We've updated the file.
                }
            }
            catch (ArgumentException)
            {
                isSuccess = false;
            }

            return isSuccess;
        }
    }
}
