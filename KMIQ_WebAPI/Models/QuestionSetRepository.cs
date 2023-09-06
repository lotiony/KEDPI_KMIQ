using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using loCommon;

namespace KMIQ.Models
{
    public class QuestionSetRepository : IRepository<QuestionSet>
    {
        List<QuestionSet> _dbContext;

        public QuestionSetRepository()
        {
            _dbContext = new List<QuestionSet>();


            using (SqlConn sql = new SqlConn())
            {
                string qry = string.Format(" EXEC [dbo].[USP_GET_QUESTIONS] ");
                DataTable rstDt = sql.SqlSelect(qry);

                if (rstDt.Rows.Count > 0)
                {
                    foreach (DataRow r in rstDt.Rows)
                    {
                        QuestionSet item = new QuestionSet()
                        {

                            Id = Convert.ToInt32(r["QID"]),
                            TypeID = Convert.ToInt32(r["TYPEID"]),
                            Number = Convert.ToInt32(r["QNUMBER"]),
                            Question = DBNull.Value.Equals(r["QUESTION"]) ? "" : r["QUESTION"].ToString(),
                            IQ_Area = DBNull.Value.Equals(r["IQ_AREA"]) ? "" : r["IQ_AREA"].ToString(),
                            Sub_Area = DBNull.Value.Equals(r["SUB_AREA"]) ? "" : r["SUB_AREA"].ToString(),
                            Classification = DBNull.Value.Equals(r["CLASSIFICATION"]) ? "" : r["CLASSIFICATION"].ToString()
                        };
                        _dbContext.Add(item);
                    }
                }
            }

        }

        public IEnumerable<QuestionSet> GetAll()
        {
            return _dbContext;
        }

        public QuestionSet FindById(int id)
        {
            return _dbContext.Where(x => x.Id == id).FirstOrDefault();
        }

        public string Update(QuestionSet entity)
        {
            throw new NotImplementedException();
        }

        public string Add(QuestionSet entity)
        {
            throw new NotImplementedException();
        }

        public string Delete(QuestionSet entity)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<QuestionSet> SelectById(int id)
        {
            try
            {
                List<QuestionSet> list = _dbContext.Where(x => x.TypeID == id).ToList();
                return list;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}