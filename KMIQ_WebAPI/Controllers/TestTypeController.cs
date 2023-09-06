using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using KMIQ.Models;

namespace KMIQ.Controllers
{
    public class TestTypeController : ApiController
    {
        IRepository<TestType> repository = new TestTypeRepository();

        public IEnumerable<TestType> GetAllTestType()
        {
            return repository.GetAll();
        }
    }
}
