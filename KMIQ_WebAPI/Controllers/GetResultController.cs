using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using KMIQ.Models;
using System.Web;
using System.Text;
using System.Collections.Specialized;
using loCommon;

namespace KMIQ.Controllers
{
    public class GetResultController : ApiController
    {
        IRepository<Result> repository = new ResultRepository();

        [HttpPost]
        [ActionName("GetResult")]
        public HttpResponseMessage PostGetResult(Result result)
        {
            if (result != null)
            {

                // 상태 텍스트에 포함되어 있는 모든 HTML 마크업을 변환합니다.
                //result.Token = HttpUtility.HtmlEncode(result.Token);
                HttpResponseMessage response = new HttpResponseMessage();

                string updateResult = repository.Update(result);

                switch (updateResult)
                {
                    case "SUCCESS":

                        //response = Request.CreateResponse(HttpStatusCode.OK);
                        //response.Content = new StringContent(updateResult, Encoding.UTF8);


                        // 301 응답을 생성합니다. 결괄페이지로 즉시 redirect 됩니다.
                        string actionUri = string.Format("{0}", result.returnUrl);
                        response = Request.CreateResponse(HttpStatusCode.Moved);
                        response.Headers.Location = new Uri(actionUri);
                        break;

                    default:
                        response = Request.CreateResponse(HttpStatusCode.OK);
                        response.Content = new StringContent(updateResult, Encoding.UTF8);
                        break;

                }


                /// 결과를 지정된 주소로 POST 전송한다.
                using (WebClient wc = new WebClient())
                {
                    try
                    {
                        string actionUri = string.Format("{0}", result.updateUrl);

                        NameValueCollection nvc = new NameValueCollection();
                        nvc.Add("result", updateResult);
                        nvc.Add("Token", result.Token);

                        wc.Encoding = UTF8Encoding.UTF8;
                        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";

                        /// POST로 값을 날리고 byte[]를 리턴받는다. response를 문자열로 변환한다.
                        Encoding encoder = UTF8Encoding.UTF8;
                        wc.UploadValues(actionUri, "POST", nvc);
                    }
                    catch (Exception)
                    {
                    }
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
