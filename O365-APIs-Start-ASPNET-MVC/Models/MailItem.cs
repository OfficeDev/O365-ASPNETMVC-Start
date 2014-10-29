// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using Microsoft.Office365.OutlookServices;
using O365_APIs_Start_ASPNET_MVC.Helpers;
using System;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace O365_APIs_Start_ASPNET_MVC.Models
{
    //Represents a Mail message in an easily consumable form by our views
    public class MailItem
    {
        public string ID;

        [DataType(DataType.MultilineText)]
        public string Body { get; set; }
        public string Recipients { get; set; }
        public string Subject { get; set; }
        public string Sender { get; set; }

        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:MM-dd-yyyy HH:mm tt}")]
        public DateTimeOffset? Received { get; set; }

        //The next two are used for implementing paging
        //Determines if this item is the first item or the last item in the collection
        public bool IsFirstItem { get; set; }
        public bool IsLastItem { get; set; }

        private MailOperations _mailOperations = new MailOperations();

        public MailItem(IMessage serverMailItem)
        {

                IsLastItem = false;
                IsFirstItem = false;

                ID = serverMailItem.Id;

                //If HTML, take text. Otherwise, use content as is
                string bodyType = serverMailItem.Body.ContentType.ToString();
                string bodyContent = serverMailItem.Body.Content;
                if (bodyType == "HTML")
                {
                    bodyContent = Regex.Replace(bodyContent, "<[^>]*>", "");
                    bodyContent = Regex.Replace(bodyContent, "\n", "");
                    bodyContent = Regex.Replace(bodyContent, "\r", "");
                }
                Body = bodyContent;

                Subject = serverMailItem.Subject;

                Recipients = _mailOperations.BuildRecipientList(serverMailItem.ToRecipients);

                if (serverMailItem.Sender != null)
                {
                    Sender = serverMailItem.Sender.EmailAddress.Address;
                }
                else
                    Sender = string.Empty; // Sometimes, mails exist as draft, and therefore haven't been sent.

                if (serverMailItem.DateTimeReceived != null)
                {
                    Received = serverMailItem.DateTimeReceived;
                }
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