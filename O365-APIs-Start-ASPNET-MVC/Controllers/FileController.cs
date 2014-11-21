using O365_APIs_Start_ASPNET_MVC.Helpers;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web.Mvc;
using model = O365_APIs_Start_ASPNET_MVC.Models;
using System.Linq;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.OpenIdConnect;
using System.Web;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace O365_APIs_Start_ASPNET_MVC.Controllers
{

    [Authorize]
    [HandleError(ExceptionType = typeof(AdalException))]
    public class FileController : Controller
    {
        private FileOperations _fileOperations = new FileOperations();
       
        private static bool _O365ServiceOperationFailed = false;  

        //Returns the user's my files
        //Implements Office 365-side paging
        // GET: /File/
        public async Task<ActionResult> Index()
        {
            ViewBag.O365ServiceOperationFailed = _O365ServiceOperationFailed;

            if (_O365ServiceOperationFailed)
            {
                _O365ServiceOperationFailed = false;
            }

            List<model.FileObject> myFiles = new List<model.FileObject>();
            try
            {
               myFiles = await _fileOperations.GetMyFilesAsync();
            }
            catch (AdalException e)
            {

                if (e.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                {

                    //This exception is thrown when either you have a stale access token, or you attempted to access a resource that you don't have permissions to access.
                    throw e;

                }

            }
            return View(myFiles);
        }

        //
        // GET: /Files/Create - This GET operation returns the Create view (Create.cshtml)
        [HttpGet]
        public ActionResult Create()
        {
            return View();
        }


        //
        // POST: /Files/Create - This POST operation creates the text file and returns the default files list view (Index.cshtml) 
        [HttpPost]
        public async Task<ActionResult> Create(FormCollection collection)
        {
            _O365ServiceOperationFailed = false;
            String newEventID = "";

            try
            {
                newEventID = await _fileOperations.CreateNewTextFileAsync();
            }
            catch (Exception)
            {
                _O365ServiceOperationFailed = true;
            }
            return RedirectToAction("Index", new { newid = newEventID });
        }




        //GET: /Files/Edit - This GET operation gets a handle on the selected file and returns the text file contents to the edit view (Edit.cshtml)
        [HttpGet]
        public async Task<ActionResult> Edit(string id)
        {
            var files = await _fileOperations.GetMyFilesAsync();
            var fileToEdit = files.Where(f => f.ID == id).SingleOrDefault();
            if (fileToEdit != null)
            {
                var results = await _fileOperations.ReadTextFileAsync(id);
                fileToEdit.FileText = results[0].ToString();
            }
            return View(fileToEdit);
        }


        // POST: /Files/Edit - This POST operation gets a handle on the selected file and saves/updates the text file contents to the file in O365
        [HttpPost]
        [ValidateInput(false)]
        public async Task<ActionResult> Edit(string id, string fileText)
        {
            _O365ServiceOperationFailed = false;

            try
            {
                await _fileOperations.UpdateTextFileAsync(id, fileText);
            }
            catch (Exception)
            {
                _O365ServiceOperationFailed = true;
            }
            return RedirectToAction("Index", new {changedid = id });
        }


        // GET: /Files/Delete - This GET operation gets a handle on the selected file or folder marked for deletion and returns the delete view (Delete.cshtml) 
        [HttpGet]
        public async Task<ActionResult> Delete(string id)
        {
            var files = await _fileOperations.GetMyFilesAsync();
            var fileToDelete = files.Where(f => f.ID == id).SingleOrDefault();
            return View(fileToDelete);
        }


        // POST: /Files/Delete - This POST operation gets a handle on the selected file or folder and deletes the it in O365
        [HttpPost]
        public async Task<ActionResult> Delete(string id, FormCollection collection)
        {
            _O365ServiceOperationFailed = false;

            try
            {
                await _fileOperations.DeleteFileOrFolderAsync(id);
            }
            catch (Exception)
            {
                _O365ServiceOperationFailed = true;
            }
            return RedirectToAction("Index");
        }
    }
}