using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Net;
using System.Text;
using System.IO;
using System.Data;
using System.Windows.Forms;
using System.Collections;

namespace KMIQ
{
    public class WebApi : IDisposable
    {
        #region public properties
        public bool IsEnabled { get; private set; }
        #endregion



        #region private common variables / properties
        bool disposed = false;
        string accessKey = "";

        string PERSONAL_REPORT_URI { get { return "http://kedpi.cafe24.com/mypage/update_report_publish.php?guid="; } }

        #endregion

        #region event

        public delegate void dCompleteSendResult(string guid, bool success);
        public event dCompleteSendResult CompleteSendResult;

        #endregion


        public WebApi()
        {
            IsEnabled = true;
        }

        ///// <summary>
        ///// API 키 발급
        ///// </summary>
        //private bool getAccessKey()
        //{
        //    bool rtn = false;
        //    accessKey = "";

        //    if (!wc.IsBusy)
        //    {
        //        try
        //        {
        //            string actionUri = string.Format("{0}?examKey={1}", REQUEST_ACCESSKEY_URI, Settings.Default.ApiKey);
        //            Stream response = wc.OpenRead(actionUri);
        //            string responseJSon = new StreamReader(response).ReadToEnd();

        //            if (responseJSon.Length > 0)
        //            {

        //                Dictionary<string, string> summaryData;
        //                summaryData = JsonConverter.DeserializeJsonUsingJavaScript<Dictionary<string, string>>(responseJSon);

        //                if (summaryData["resultCode"] == "0")
        //                {
        //                    accessKey = summaryData["accessKey"];
        //                    rtn = true;
        //                }
        //                else
        //                {
        //                    accessKey = "";
        //                    MessageBox.Show(string.Format("LMS 서버와 연동을 위한 API Key 발급에 실패했습니다.{0}- 에러코드 : {1}{0}- 에러내용 : {2}", Environment.NewLine, summaryData["resultCode"].ToString(), summaryData["resultMessage"].ToString()));
        //                }
        //            }
        //            else
        //            {
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            loFunctions.LogWrite("[getAccessKey] on Error : " + ex.Message);
        //            MessageBox.Show("LMS 서버와 연동을 위한 API Key 발급에 실패했습니다." + Environment.NewLine + ex.Message);
        //        }

        //    }

        //    return rtn;
        //}


        /// <summary>
        /// 사용자 답안을 전송하고 결과값을 리턴한다.
        /// </summary>
        public string SendResult(string guid)
        {
            string rtn = "";

            try
            {
                MyWebClient wc = new MyWebClient();
                wc.DownloadStringCompleted += Wc_DownloadStringCompleted;

                string actionUri = this.PERSONAL_REPORT_URI + guid;
                wc.Headers.Clear();
                //wc.Headers.Add(string.Format("X-Auth-AccessKey:{0}", accessKey));

                //NameValueCollection nvc = new NameValueCollection
                //{
                //    { "personal_id", personal_id },
                //    { "paper_id", papaer_id },
                //    { "data", omrResult }
                //};

                //wc.Encoding = UTF8Encoding.UTF8;
                //wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";

                ///// 호출콜을 로그에 쓴다.
                //loFunctions.LogWrite($"[[[ WEB API Logging ]]] ==> {actionUri}?personal_id={personal_id}&paper_id={papaer_id}&data={omrResult}");

                /// GET으로 값을 날리고 byte[]를 리턴받는다. response를 문자열로 변환한다.
                Encoding encoder = UTF8Encoding.UTF8;
                wc.DownloadStringAsync(new Uri(actionUri), guid);

            }
            catch (WebException ex)
            {
                //var response = ((HttpWebResponse)ex.Response).StatusCode;
                //switch (response)
                //{
                //    case HttpStatusCode.BadRequest:
                //        break;
                //    case HttpStatusCode.Conflict:
                //        break;
                //    case HttpStatusCode.Forbidden:
                //        break;
                //}
                rtn = ((HttpWebResponse)ex.Response).StatusCode.ToString();
            }

            return rtn;
        }

        private void Wc_DownloadStringCompleted(object sender, DownloadStringCompletedEventArgs e)
        {
            bool success = false;
            if (e.Result.Contains("OK")) success = true;

            CompleteSendResult(e.UserState.ToString(), success);
        }






        /// <summary>
        /// 개인리포트 브라우저를 열어준다.
        /// </summary>
        public void GotoPersonalReport(string personal_id)
        {
            try
            {
                string actionUri = this.PERSONAL_REPORT_URI;

                string strPostData = $"id={personal_id}";
                byte[] postData = Encoding.Default.GetBytes(strPostData);

                

            }
            catch (WebException ex)
            {

            }
            
        }




        private static DataTable DictionaryToDataTable(List<Dictionary<string, string>> list)
        {
            DataTable result = new DataTable();
            if (list.Count == 0)
                return result;

            result.Columns.AddRange(
                list.First().Select(r => new DataColumn(r.Key)).ToArray()
            );

            list.ForEach(r => result.Rows.Add(r.Select(c => c.Value).Cast<object>().ToArray()));

            return result;
        }




        private static DataTable ArrayListToDataTable(ArrayList list)
        {
            DataTable result = new DataTable();
            if (list.Count == 0)
                return result;

            Dictionary<string, object> firstItem = (Dictionary<string, object>)list[0];

            result.Columns.AddRange(
                firstItem.Select(r => new DataColumn(r.Key, typeof(string))).ToArray()
            );

            foreach (var items in list)
            {
                Dictionary<string, object> it = (Dictionary<string, object>)items;

                result.Rows.Add(it.Values.ToArray());
            }

            return result;
        }




        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            disposed = true;
        }

        #endregion


    }

    public class MyWebClient : WebClient
    {
        protected override WebRequest GetWebRequest(Uri address)
        {
            WebRequest request = base.GetWebRequest(address);
            request.Timeout = 1200000; // 1200 sec = 20 min
            return request;
        }
    }

}
