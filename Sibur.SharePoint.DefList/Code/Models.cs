using System.Collections.Generic;


namespace Sibur.SharePoint.DefList
{
    public class Divisions
    {
        public IList<Division> divisions { get; set; }
    }

    public class Division
    {
        public string id { get; set; }
        public string title { get; set; }
        public IList<Work> works { get; set; }
    }


    public class Work
    {
        public string id { get; set; }
        public string title { get; set; }
        public string units { get; set; }
        public string count { get; set; }
        public string mtp { get; set; }
        public string comment { get; set; }
    }

}
