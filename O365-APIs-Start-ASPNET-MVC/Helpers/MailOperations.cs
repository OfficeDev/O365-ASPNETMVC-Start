// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office365.OutlookServices;
using model = O365_APIs_Start_ASPNET_MVC.Models;
using System.Threading.Tasks;

namespace O365_APIs_Start_ASPNET_MVC.Helpers
{
    /// <summary>
    /// Contains methods for making requests against Office 365 email.
    /// </summary>
    internal class MailOperations
    {
        /// <summary>
        /// Fetches email from user's Inbox.
        internal async Task<List<model.MailItem>> GetEmailMessages(int pageNo, int pageSize)
        {
            // Make sure we have a reference to the Outlook Services client
            var outlookServicesClient = await AuthenticationHelper.EnsureOutlookServicesClientCreatedAsync("Mail");

            List<model.MailItem> returnResults = new List<model.MailItem>();

            var mailResults = await (from i in outlookServicesClient.Me.Folders.GetById("Inbox").Messages
                                     orderby i.DateTimeReceived descending
                                     select i).Skip((pageNo - 1) * pageSize).Take(pageSize).ExecuteAsync();

            var mailMessages = mailResults.CurrentPage;

            foreach (IMessage serverMailItem in mailMessages)
            {
                model.MailItem modelMailItem = new model.MailItem(serverMailItem);
                if ((!mailResults.MorePagesAvailable))
                {
                    if (modelMailItem.ID == mailMessages.Last<IMessage>().Id)
                    {
                        modelMailItem.IsLastItem = true;

                    }
                }
                if ((modelMailItem.ID == mailMessages.First<IMessage>().Id) && pageNo == 1)
                {
                    modelMailItem.IsFirstItem = true;
                }

                returnResults.Add(modelMailItem);
            }

            return returnResults;
        }

        /// <summary>
        /// Compose and send a new email.
        /// </summary>
        /// <param name="subject">string. The subject line of the email.</param>
        /// <param name="bodyContent">string. The body of the email.</param>
        /// <param name="recipients">string. A semi-colon separated list of email addresses.</param>
        /// <returns></returns>
        internal async Task<String> ComposeAndSendMailAsync(string subject,
                                                            string bodyContent,
                                                            string recipients)
        {
            // The identifier of the composed and sent message.
            string newMessageId = string.Empty;

            // Prepare the recipient liost
            var toRecipients = new List<Recipient>();
            string[] splitter = { ";" };
            var splitRecipientsString = recipients.Split(splitter, StringSplitOptions.RemoveEmptyEntries);
            foreach (string recipient in splitRecipientsString)
            {
                toRecipients.Add(new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = recipient.Trim(),
                        Name = recipient.Trim(),
                    },
                });
            }

            // Prepare the draft message.
            var draft = new Message
            {
                Subject = subject,
                Body = new ItemBody 
                           { 
                                ContentType = BodyType.Text,
                                Content = bodyContent
                           },
                ToRecipients = toRecipients,
            };

            try
            {
                // Make sure we have a reference to the Outlook Services client.
                var outlookServicesClient = await AuthenticationHelper.EnsureOutlookServicesClientCreatedAsync("Mail");

                // To send a message without saving to Sent Items, specify false for  
                // the SavetoSentItems parameter. This method is useful only when you 
                // don't need a handle (identifier) on the item saved in 
                // the sent items folder.
                // await outlookServicesClient.Me.SendMailAsync(draft, true);


                // Creates the draft message in the drafts folder. 
                // This results in a call to the service. 
                // Returns full item but unfortunately you dont have access to it.
                await outlookServicesClient.Me.Folders.GetById("Drafts").Messages.AddMessageAsync(draft);

                // Gets the full draft message, including the identifier needed to issue a send mail request.
                // This results in a call to the service. 
                IMessage updatedDraft = await outlookServicesClient.Me.Folders.GetById("Drafts").Messages.GetById(draft.Id).ExecuteAsync();

                // Issues a send command so that the draft mail is sent to the recipient list.
                // This results in a call to the service. 
                await outlookServicesClient.Me.Folders.GetById("Drafts").Messages.GetById(updatedDraft.Id).SendAsync();

                newMessageId = draft.Id;
            }
            catch (Exception e)
            {
                throw new Exception("We could not send the message: " + e.Message);
            }
            return newMessageId;
        }

        /// <summary>
        /// Removes an event from the user's default calendar.
        /// </summary>
        /// <param name="selectedEventId">string. The unique Id of the event to delete.</param>
        /// <returns></returns>
        internal async Task<IMessage> DeleteMailItemAsync(string selectedEventId)
        {
            IMessage thisMailItem = null;
            try
            {
                // Make sure we have a reference to the Outlook Services client
                var outlookServicesClient = await AuthenticationHelper.EnsureOutlookServicesClientCreatedAsync("Mail");

                // Get the event to be removed from the Exchange service. This results in a call to the service.
                thisMailItem = await outlookServicesClient.Me.Folders.GetById("Inbox").Messages.GetById(selectedEventId).ExecuteAsync();

                // Delete the event. This results in a call to the service.
                await thisMailItem.DeleteAsync(false);
            }
            catch (Exception)
            {
                throw new Exception("The message could not be deleted in Outlook Services.");
            }
            return thisMailItem;
        }

        internal string BuildRecipientList(IList<Recipient> recipientList)
        {
            StringBuilder recipientListBuilder = new StringBuilder();
            foreach (Recipient recipient in recipientList)
            {
                if (recipientListBuilder.Length == 0)
                {
                    recipientListBuilder.Append(recipient.EmailAddress.Address);
                }
                else
                {
                    recipientListBuilder.Append(";" + recipient.EmailAddress.Address);
                }
            }

            return recipientListBuilder.ToString();
        }

        internal async Task<model.MailItem> GetMailItemByIDsAsync(string id)
        {
            // Make sure we have a reference to the Outlook Services client
            var outlookServicesClient = await AuthenticationHelper.EnsureOutlookServicesClientCreatedAsync("Mail");
            IMessage thisMailItem = null;

            // This results in a call to the service.
            var thisMailFetcher = outlookServicesClient.Me.Folders.GetById("Inbox").Messages.GetById(id);
            thisMailItem = await thisMailFetcher.ExecuteAsync();
            model.MailItem modelMailMessage = new model.MailItem(thisMailItem);
            return modelMailMessage;
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