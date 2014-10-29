// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using Microsoft.Office365.OutlookServices;
using O365_APIs_Start_ASPNET_MVC.Helpers;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web.Mvc;
using model = O365_APIs_Start_ASPNET_MVC.Models;

namespace O365_APIs_Start_ASPNET_MVC.Controllers
{
    //Read, , create, edit, and delete contacts. 

    [Authorize]
    public class ContactController : Controller
    {
        private ContactOperations _contactOperations = new ContactOperations();
        private AuthenticationHelper _authHelper = new AuthenticationHelper();

        //Returns the user's contacts
        //Implements Office 365-side paging
        // GET: /Contact/
        public async Task<ActionResult> Index(int? page)
        {
            var pageNumber = page ?? 1;
            if (page < 1)
            {

                pageNumber = 1;

            }
            //Number of events displayed on one page. Edit pageSize if you like
            int pageSize = 10;

            List<model.ContactItem> contacts = await _contactOperations.GetContactsPageAsync(pageNumber, pageSize);

            //Store these in the ViewBag so you can use them in the Index view
            ViewBag.Page = pageNumber;
            ViewBag.NextPage = pageNumber + 1;
            ViewBag.PrevPage = pageNumber - 1;
            ViewBag.LastPage = false;

            if ((contacts != null) && (contacts.Count == 0))
            {
                ViewBag.LastPage = true;
            }
            return View(contacts);
        }
        //
        // GET: /Contact/Create
        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /Contact/Create
        [HttpPost]
        public async Task<ActionResult> Create(FormCollection collection)
        {
            String newEventID = await _contactOperations.AddContactItemAsync(collection["FileAs"], 
                                                                             collection["GivenName"], 
                                                                             collection["Surname"], 
                                                                             collection["JobTitle"], 
                                                                             collection["Email"], 
                                                                             collection["MobilePhone"], 
                                                                             collection["BusinessPhone"]);
            return RedirectToAction("Index", new { newid = newEventID });
        }

        //
        // GET: /Contact/Edit/5
        public async Task<ActionResult> Edit(string id, int page)
        {
            model.ContactItem contactToUpdate = await _contactOperations.GetContactByIDsAsync(id);
            return View(contactToUpdate);
        }

        //
        // POST: /Contact/Edit/5
        [HttpPost]
        public async Task<ActionResult> Edit(string id, int page, FormCollection collection)
        {
            IContact updatedContact = await _contactOperations.UpdateContactItemAsync(id, 
                                                                                      collection["FileAs"], 
                                                                                      collection["GivenName"], 
                                                                                      collection["Surname"], 
                                                                                      collection["JobTitle"], 
                                                                                      collection["Email"], 
                                                                                      collection["MobilePhone"], 
                                                                                      collection["BusinessPhone"]);
            return RedirectToAction("Index", new { page, changedid = id });
        }

        //
        // GET: /Contact/Delete/5
        public async Task<ActionResult> Delete(string id)
        {
            model.ContactItem contactToDelete = await _contactOperations.GetContactByIDsAsync(id);
            return View(contactToDelete);
        }

        //
        // POST: /Contact/Delete/5
        [HttpPost]
        public async Task<ActionResult> Delete(string id, FormCollection collection)
        {
            bool success = await _contactOperations.DeleteContactAsync(id);
            return RedirectToAction("Index");
        }
    }
}
//*********************************************************  
//  
//O365 APIs Starter Project for ASPNET MVC, https://github.com/OfficeDev/Office-365-APIs-Starter-Project-for-ASPNETMVC
// 
//Copyright (c) Microsoft Corporation 
//All rights reserved.  
// 
//MIT License: 
// 
//Permission is hereby granted, free of charge, to any person obtaining 
//a copy of this software and associated documentation files (the 
//""Software""), to deal in the Software without restriction, including 
//without limitation the rights to use, copy, modify, merge, publish, 
//distribute, sublicense, and/or sell copies of the Software, and to 
//permit persons to whom the Software is furnished to do so, subject to 
//the following conditions: 
// 
//The above copyright notice and this permission notice shall be 
//included in all copies or substantial portions of the Software. 
// 
//THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
//EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
//MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
//NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
//LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
//OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
//WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. 
//  
//********************************************************* 
