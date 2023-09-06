using KMIQ.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web;
using System.Web.Http;

namespace KMIQ.Controllers
{
    public class GetReportFileController : ApiController
    {
        ReportFileRepository repository = new ReportFileRepository();

        [HttpGet]
        public HttpResponseMessage GetReport(string id)
        {
            if (id != null && id != "")
            {

                // 상태 텍스트에 포함되어 있는 모든 HTML 마크업을 변환합니다.
                //result.Token = HttpUtility.HtmlEncode(result.Token);
                HttpResponseMessage response = new HttpResponseMessage();

                ReportFile result = repository.GetFileByToken(id);

                
                switch (result.returnMsg)
                {
                    case "OK":

                        response = Request.CreateResponse(HttpStatusCode.OK);
                        //response.Content = new StringContent(updateResult, Encoding.UTF8);

                        string fileName = result.fileName;
                        string strFileName = "";
                        if (HttpContext.Current.Request.UserAgent.IndexOf("NT 5.0") >= 0)
                        {
                            strFileName = HttpUtility.UrlEncode(fileName);
                        }
                        else
                        {
                            strFileName = HttpUtility.UrlEncode(fileName, new UTF8Encoding()).Replace("+", "%20");
                        }

                        Stream st = new MemoryStream(result.fileData);
                        response.Content = new StreamContent(st);
                        response.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
                        response.Content.Headers.ContentDisposition.FileName = strFileName;
                        response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
                        
                        break;

                    default:
                        string alertMsg = string.Format("<script>alert('{0}'); window.open('about: blank','_self').close();</script>", result.returnMsg);
                        response = Request.CreateResponse(HttpStatusCode.OK);
                        response.Content = new StringContent(alertMsg, Encoding.UTF8, "text/html");
                        break;

                }

                return response;
            }
            else
            {
                return Request.CreateResponse(HttpStatusCode.BadRequest);
            }
        }


    }
}
