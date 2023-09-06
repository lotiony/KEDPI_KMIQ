using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using loCommon;

namespace KMIQ.Models
{
    public class TestTypeRepository : IRepository<TestType>
    {
        List<TestType> _dbContext;

        public TestTypeRepository()
        {
            _dbContext = new List<TestType>();


            using (SqlConn sql = new SqlConn())
            {
                string qry = string.Format(" SELECT * FROM T_TEST_TYPE ");
                DataTable rstDt = sql.SqlSelect(qry);

                if (rstDt.Rows.Count > 0)
                {
                    foreach (DataRow r in rstDt.Rows)
                    {
                        TestType item = new TestType()
                        {
                            Id = Convert.ToInt32(r["TYPEID"]),
                            TypeGrade = r["TYPE_GRADE"].ToString(),
                            TypeMark = r["TYPE_MARK"].ToString()

                        };
                        _dbContext.Add(item);
                    }
                }
            }

        }

        public IEnumerable<TestType> GetAll()
        {
            return _dbContext;
        }

        public TestType FindById(int id)
        {
            return _dbContext.Where(x => x.Id == id).FirstOrDefault();
        }

        public string Update(TestType entity)
        {
            throw new NotImplementedException();
        }

        public string Add(TestType entity)
        {
            throw new NotImplementedException();
        }

        public string Delete(TestType entity)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<TestType> SelectById(int id)
        {
            throw new NotImplementedException();
        }
    }
}