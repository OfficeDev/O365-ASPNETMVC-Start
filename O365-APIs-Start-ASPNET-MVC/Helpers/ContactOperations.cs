// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using Microsoft.Office365.OutlookServices;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using model = O365_APIs_Start_ASPNET_MVC.Models;

namespace O365_APIs_Start_ASPNET_MVC.Helpers
{
    /// <summary>
    /// Contains methods for making requests against Office 365 contacts.
    /// </summary>
    internal class ContactOperations
    {
        internal async Task<List<model.ContactItem>> GetContactsPageAsync(int pageNo, int pageSize)
        {
            try
            {
                // Get exchangeclient
                var outlookServicesClient = await AuthenticationHelper.EnsureOutlookServicesClientCreatedAsync("Contacts");

                List<model.ContactItem> returnResults = new List<model.ContactItem>();

                // Get contacts
                var contactsResults = await (from i in outlookServicesClient.Me.Contacts
                                             orderby i.FileAs
                                             select i).Skip((pageNo - 1) * pageSize).Take(pageSize).ExecuteAsync();
                var contacts = contactsResults.CurrentPage;

                foreach (IContact serverContactItem in contacts)
                {

                    model.ContactItem modelContact = new model.ContactItem(serverContactItem);
                    if ((!contactsResults.MorePagesAvailable))
                    {
                        if (modelContact.ID == contacts.Last<IContact>().Id)
                        {
                            modelContact.IsLastItem = true;

                        }
                    }
                    if ((modelContact.ID == contacts.First<IContact>().Id) && pageNo == 1)
                    {
                        modelContact.IsFirstItem = true;
                    }

                    returnResults.Add(modelContact);
                }


                return returnResults;
            }
            catch { return null; }
        }

        /// <summary>
        /// Adds a new contact.
        /// </summary>
        internal async Task<string> AddContactItemAsync(

            string fileAs,
            string givenName,
            string surname,
            string jobTitle,
            string email,
            string workPhone,
            string mobilePhone
            )
        {
            Contact newContact = new Contact
            {
                FileAs = fileAs,
                GivenName = givenName,
                Surname = surname,
                JobTitle = jobTitle,
                MobilePhone1 = mobilePhone
            };

            newContact.BusinessPhones.Add(workPhone);


            // Note: Setting EmailAddress1 to a null or empty string will throw an exception that
            // states the email address is invalid and the contact cannot be added.
            // Setting EmailAddress1 to a string that does not resemble an email address will not
            // cause an exception to be thrown, but the value is not stored in EmailAddress1.
            if (!string.IsNullOrEmpty(email))
                newContact.EmailAddresses.Add(new EmailAddress() { Address = email });

            try
            {
                // Make sure we have a reference to the Outlook Services client
                var outlookServicesClient = await AuthenticationHelper.EnsureOutlookServicesClientCreatedAsync("Contacts");

                // This results in a call to the service.
                await outlookServicesClient.Me.Contacts.AddContactAsync(newContact);
                return newContact.Id;
            }
            catch { return null; }
        }

        /// <summary>
        /// Updates an existing contact.
        /// </summary>
        internal async Task<IContact> UpdateContactItemAsync(

            string selectedContactId,
            string fileAs,
            string givenName,
            string surname,
            string jobTitle,
            string email,
            string workPhone,
            string mobilePhone
           )
        {

            try
            {
                // Make sure we have a reference to the Outlook Services client
                var outlookServicesClient = await AuthenticationHelper.EnsureOutlookServicesClientCreatedAsync("Contacts");

                var contactToUpdate = await outlookServicesClient.Me.Contacts[selectedContactId].ExecuteAsync();

                contactToUpdate.FileAs = fileAs;
                contactToUpdate.GivenName = givenName;
                contactToUpdate.Surname = surname;
                contactToUpdate.JobTitle = jobTitle;

                contactToUpdate.MobilePhone1 = mobilePhone;

                // Note: Setting EmailAddress1 to a null or empty string will throw an exception that
                // states the email address is invalid and the contact cannot be added.
                // Setting EmailAddress1 to a string that does not resemble an email address will not
                // cause an exception to be thrown, but the value is not stored in EmailAddress1.

                //if (!string.IsNullOrEmpty(email))
                //    contactToUpdate.EmailAddress1 = email;

                // Update the contact in Exchange
                await contactToUpdate.UpdateAsync();

                return contactToUpdate;

                // A note about Batch Updating
                // You can save multiple updates on the client and save them all at once (batch) by 
                // implementing the following pattern:
                // 1. Call UpdateAsync(true) for each contact you want to update. Setting the parameter dontSave to true 
                //    means that the updates are registered locally on the client, but won't be posted to the server.
                // 2. Call exchangeClient.Context.SaveChangesAsync() to post all contact updates you have saved locally  
                //    using the preceding UpdateAsync(true) call to the server, i.e., the user's Office 365 contacts list.
            }
            catch { return null; }
        }

        /// <summary>
        /// Deletes a contact.
        /// </summary>
        /// <param name="contactId">The contact to delete.</param>
        /// <returns>True if deleted;Otherwise, false.</returns>
        internal async Task<bool> DeleteContactAsync(string contactId)
        {
            try
            {
                // Make sure we have a reference to the Outlook Services client
                var outlookServicesClient = await AuthenticationHelper.EnsureOutlookServicesClientCreatedAsync("Contacts");

                var contactToDelete = await outlookServicesClient.Me.Contacts[contactId].ExecuteAsync();

                await contactToDelete.DeleteAsync();

                return true;
            }
            catch { return false; }
        }

        internal async Task<model.ContactItem> GetContactByIDsAsync(string id)
        {
            // Make sure we have a reference to the Outlook Services client
            var outlookServicesClient = await AuthenticationHelper.EnsureOutlookServicesClientCreatedAsync("Contacts");

            // This results in a call to the service.
            var thisContactFetcher = outlookServicesClient.Me.Contacts.GetById(id);
            var thisContact = await thisContactFetcher.ExecuteAsync();
            model.ContactItem modelContact = new model.ContactItem(thisContact);
            return modelContact;
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