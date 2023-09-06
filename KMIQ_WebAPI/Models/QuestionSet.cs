using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace KMIQ.Models
{
    public class QuestionSet : IEntity
    {
        public int TypeID { get; set; }
        public int Number { get; set; }
        public string Question { get; set; }
        public string IQ_Area { get; set; }
        public string Sub_Area { get; set; }
        public string Classification { get; set; }
    }
}