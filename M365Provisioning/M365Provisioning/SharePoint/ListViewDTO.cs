﻿using System;
using Microsoft.SharePoint.Client;

namespace M365Provisioning.SharePoint
{
    public class ListViewDto
    {
        public string ListName { get; set; }
        public string ViewName { get; set; }
        public bool DefaultView { get; set; } = false;
        public ViewFieldCollection ViewFields {  get; set; } 
        public uint RowLimit { get; set; }
        public string ListScope { get; set; }
        public string JsonFormatterFile { get; set; } 

        public ListViewDto(string listName,string viewName,bool defaultView, ViewFieldCollection viewFields,uint rowLimit,string scope, string jsonFormatterFile) 
        {
            ListName = listName;
            ViewName = viewName;
            DefaultView = defaultView;
            ViewFields = viewFields;
            RowLimit = rowLimit;
            ListScope = scope;
            JsonFormatterFile = jsonFormatterFile;
        }
    }
}