using System;


namespace Sibur.SharePoint.DefList
{
    public class ResponseModel
    {
        public Guid SiteId { set; get; }
        public Guid WebId { set; get; }
        public string SiteUrl { set; get; }
        public string WebUrl { set; get; }
        public string DefListUrl { set; get; }
        public string ItemUrl { set; get; }
        public string JsonValue { set; get; }
        public bool FileExist { set; get; }
        public int ArrayLength { set; get; }
        public int JsonCount { set; get; }
        public string Message { set; get; }

    }

}
