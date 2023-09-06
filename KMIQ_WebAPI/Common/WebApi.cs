using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Net;
using System.Text;
using KMIQ.Properties;
using System.IO;
using System.Threading;
using System.Data;
using System.Collections;

namespace loCommon
{
    public class WebApi : IDisposable
    {
        #region public properties
        public bool IsEnabled { get; private set; }
        #endregion



        #region private common variables / properties
        bool disposed = false;
        WebClient wc;

        string REQUEST_MEMBERINFO_URI { get { return string.Format("{0}json/member.html", Settings.Default.ApiHost); } }

        #endregion


        public WebApi()
        {
            wc = new WebClient();

            /// 초기화시 기본 시험 데이터를 받아둔다.
            //IsEnabled = getAccessKey();
            IsEnabled = true;
        }

        /// <summary>
        /// 웹에서 token을 이용해 사용자 정보 확인
        /// </summary>
        public Dictionary<string, string> getUserInfo(string token)
        {
            Dictionary<string, string> rtn = new Dictionary<string, string>();

            if (!wc.IsBusy)
            {
                try
                {
                    string actionUri = string.Format("{0}?Token={1}", REQUEST_MEMBERINFO_URI, token);
                    //wc.Headers.Clear();
                    //wc.Headers.Add(string.Format("X-Auth-AccessKey:{0}", accessKey));

                    Stream response = wc.OpenRead(actionUri);
                    string responseStr = new StreamReader(response, Encoding.GetEncoding(51949)).ReadToEnd();

                    if (responseStr.Length > 0)
                    {
                        List<string> splitItem = responseStr.Replace("\r\n","").Trim().Split('|').ToList();

                        if (splitItem.Count > 0)
                        {
                            foreach (string item in splitItem)
                            {
                                string[] items = item.Split('=');
                                rtn.Add(items[0], items[1]);
                            }
                        }
                        else
                        {
                        }
                    }
                }
                catch (Exception ex)
                {
                    loFunctions.LogWrite("[getExamList] on Error : " + ex.Message, ex);
                }
            }

            return rtn;
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
                    wc.Dispose();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            disposed = true;
        }

        #endregion


    }
}
