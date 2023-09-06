using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace KMIQ.Models
{
    public class TestType : IEntity
    {
        public string TypeGrade { get; set; }
        public string TypeMark { get; set; }
    }
}