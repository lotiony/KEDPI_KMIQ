using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace KMIQ.Models
{
    public interface IRepository<T> where T : IEntity
    {
        IEnumerable<T> GetAll();
        string Add(T entity);
        string Delete(T entity);
        string Update(T entity);
        T FindById(int id);
        IEnumerable<T> SelectById(int id);
    }
}