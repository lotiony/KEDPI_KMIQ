using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using KMIQ.Models;

namespace KMIQ.Controllers
{
    public class QuestionSetController : ApiController
    {
        IRepository<QuestionSet> repository = new QuestionSetRepository();

        public IEnumerable<QuestionSet> GetAllQuestionSet()
        {
            return repository.GetAll();
        }

        public IEnumerable<QuestionSet> GetQuestionSetById(int id)
        {
            var result = repository.SelectById(id);
            if (result == null)
                throw new HttpResponseException(HttpStatusCode.NotFound);

            return result;
        }

    }
}
