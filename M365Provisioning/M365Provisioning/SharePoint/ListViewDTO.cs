﻿using System;

namespace M365Provisioning.SharePoint
{
    public class ListViewDto
    {
        public string ListName { get; set; } = String.Empty;
        public string ViewName { get; set; } = string.Empty;
        public bool DefaultView { get; set; } = false;
        public string ViewFields {  get; set; } = String.Empty;
        public uint RowLimit { get; set; }
        public string ListScope { get; set; } = string.Empty;
        public string JsonFormatterFile { get; set; } = string.Empty;

        public ListViewDto() { }
        public ListViewDto(string listName,string viewName,bool defaultView,string viewFields,uint rowLimit,string scope, string jsonFormatterFile) 
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
