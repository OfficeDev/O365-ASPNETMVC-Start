// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using Microsoft.Office365.OutlookServices;
using O365_APIs_Start_ASPNET_MVC.Helpers;
using System;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace O365_APIs_Start_ASPNET_MVC.Models
{
    //Represents a Calendar event in an easily consumable form by our views
    public class CalendarEvent
    {
        public string ID;

        public string Subject { get; set; }

        public string Location { get; set; }

        [Required]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:MM-dd-yyyy HH:mm tt}")]
        public DateTimeOffset StartDate { get; set; }
        
        [Required]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:MM-dd-yyyy HH:mm tt}")]
        public DateTimeOffset EndDate { get; set; }

        [DataType(DataType.MultilineText)]
        public string Body { get; set; }

        public string Attendees { get; set; }

        // These variables indicate whether an event is the first or last event in the result set. 
        public bool IsFirstItem { get; set; }
        public bool IsLastItem { get; set; }

        private CalendarOperations _calenderOperations = new CalendarOperations();

        public CalendarEvent(IEvent serverEvent)
        {
            IsLastItem = false;
            IsFirstItem = false;

            string bodyContent = string.Empty;
            if (serverEvent.Body != null)
                bodyContent = serverEvent.Body.Content;

            ID = serverEvent.Id;
            Subject = serverEvent.Subject;
            Location = serverEvent.Location.DisplayName;
            StartDate = (DateTimeOffset)serverEvent.Start.Value.ToLocalTime();
            EndDate = (DateTimeOffset)serverEvent.End.Value.ToLocalTime();


            // Remove HTML tags if the body is returned as HTML.
            string bodyType = serverEvent.Body.ContentType.ToString();
            if (bodyType == "HTML")
            {
                bodyContent = Regex.Replace(bodyContent, "<[^>]*>", "");
                bodyContent = Regex.Replace(bodyContent, "\n", "");
                bodyContent = Regex.Replace(bodyContent, "\r", "");
            }
            Body = bodyContent;
            Attendees = _calenderOperations.BuildAttendeeList(serverEvent.Attendees);
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