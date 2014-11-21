using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office365.SharePoint.FileServices;
using Microsoft.Office365.SharePoint;
using System.ComponentModel.DataAnnotations;

namespace O365_APIs_Start_ASPNET_MVC.Models
{
    public class FileObject
    {
        public string Name;
        public string DisplayName;
        public string ID;

        [DataType(DataType.MultilineText)]
        public string FileText { get; set; }

        [DataType(DataType.MultilineText)]
        public string UpdatedText { get; set; }


        public FileObject(IItem fileItem)
        {

            ID = fileItem.Id;

            Name = fileItem.Name;

            DisplayName = (fileItem is Folder) ? "Folder" : "File";
           
        }
    }
}