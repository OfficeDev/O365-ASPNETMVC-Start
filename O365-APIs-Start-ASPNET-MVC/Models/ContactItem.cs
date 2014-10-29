// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using Microsoft.Office365.OutlookServices;

namespace O365_APIs_Start_ASPNET_MVC.Models
{
    //Represents a contact item in an easily consumable form by our views
    public class ContactItem
    {
        public string ID;

        public string FileAs { get; set; }
        public string GivenName { get; set; }

        public string Surname { get; set; }

        public string JobTitle { get; set; }

        public string Email { get; set; }

        public string MobilePhone { get; set; }

        public string BusinessPhone { get; set; }

        //The next two are used for implementing paging
        //Determines if this item is the first item or the last item in the collection
        public bool IsFirstItem { get; set; }
        public bool IsLastItem { get; set; }

        public ContactItem(IContact serverContactItem)
        {
                IsLastItem = false;
                IsFirstItem = false;

                ID = serverContactItem.Id;
                FileAs = serverContactItem.FileAs;
                GivenName = serverContactItem.GivenName;
                Surname = serverContactItem.Surname;
                JobTitle = serverContactItem.JobTitle;
                MobilePhone = serverContactItem.MobilePhone1;
                if (serverContactItem.BusinessPhones[0] != null)
                {
                    BusinessPhone = serverContactItem.BusinessPhones[0];
                }
                if (serverContactItem.EmailAddresses[0] != null)
                {
                    Email = serverContactItem.EmailAddresses[0].Address;
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

