using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;


namespace KMIQ.Models
{
    public class Result : IEntity
    {
        public string ID { get; set; }

        public int TypeId { get; set; }

        public string Token { get; set; }

        public string ResultStr { get; set; }

        public string uName { get; set; }
        public string uBirth { get; set; }
        public string uEmail { get; set; }
        public string uTel { get; set; }   


        public string returnUrl { get; set; }
        public string updateUrl { get; set; }
    }
}