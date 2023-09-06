using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace KMIQ.Models
{
    public class ReportFile : IEntity
    {
        public bool isEnabled { get; set; }
        public string fileName { get; set; }
        public byte[] fileData { get; set; }
        public string returnMsg { get; set; }
    }
}