using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Microsoft.VisualBasic.Logging;

namespace loCommon
{
	public static class loFunctions
	{
		[DllImport("ole32.dll")]
		public static extern void CoUninitialize();

		[DllImport("user32.dll")]
		private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

		///******************************************************************************************************************************************************
		/// 오브젝트 리소스 해제
		///******************************************************************************************************************************************************
		public static void releaseObject(object obj)
		{
			try
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
				obj = null;
			}
			catch (Exception ex)
			{
				obj = null;
				throw ex;
			}
		}

    

		public static string getMyIP()
		{
			string myIP = "0.0.0.0";

			try
			{
				using (WebClient wc = new WebClient())
				{
                    string RequestHtml = wc.DownloadString("Http://www.youhost.co.kr/ip.php");

                    //MessageBox.Show(RequestHtml)
                    string[] separ = new string[] { "접근하신 IP 주소는", " 입니다" };
                    string[] result = RequestHtml.Split(separ, StringSplitOptions.RemoveEmptyEntries);

                    myIP = result[1].Trim();
                    //myIP = RequestHtml.Substring("접근하신 IP 주소는 ".Length, RequestHtml.IndexOf(" 입니다.") - "접근하신 IP 주소는 ".Length).Trim();
                    //MessageBox.Show(USER_IP)
				}
			}
			catch { }
			return myIP;
		}




		//******************************************************************************************************************************************************
		// 로그를 기록함
		//******************************************************************************************************************************************************
		private static  Log log = new Log();
		public static void LogWrite(string log_text, Exception ex)
		{
			// 로그 기록 폴더        
			log.DefaultFileLogWriter.CustomLocation = AppDomain.CurrentDomain.BaseDirectory + "\\Log";
            if (!Directory.Exists(log.DefaultFileLogWriter.CustomLocation)) Directory.CreateDirectory(log.DefaultFileLogWriter.CustomLocation);

			// 로그 파일 명(프로그램명_날짜)        
			log.DefaultFileLogWriter.BaseFileName = "ErrorLog_" + DateTime.Now.ToString("yyyy-MM-dd");
			// 로그 내용 기록        
			log.WriteEntry( String.Format( DateTime.Now.ToString(), "yyyy-MM-dd HH:mm:ss") + "  ===  " + log_text, TraceEventType.Information);
			// 로그 기록 닫기        
			log.DefaultFileLogWriter.Close();
		}




		//******************************************************************************************************************************************************
		// 경과시간 측정도구
		//******************************************************************************************************************************************************
		private static Stopwatch sw;
		private static Stopwatch sw2;
		public static void StartBenchmark()
		{
			sw = new Stopwatch();
			sw2 = new Stopwatch();
			sw.Start();
			sw2.Start();
		}

		public static string ElapsedBenchMark()
		{
			TimeSpan ts = sw.Elapsed;
			return String.Format("{0:00}:{1:00}:{2:00}.{3:000}", 	ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds);
		}
		public static string ElapsedBenchMark2()
		{
			TimeSpan ts = sw2.Elapsed;
			sw2.Restart();
			return String.Format("{0:00}:{1:00}:{2:00}.{3:000}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds);
		}

		public static void StopBenchmark()
		{
			try
			{
				sw.Stop();
				sw = null;
				sw2.Stop();
				sw2 = null;
			}
			catch { }
		}



        public static DataSet getDataSetFromDB(string qry)
        {
            DataSet rstDs = null;

            using (SqlConn sql = new SqlConn())
            {
                try
                {
                    rstDs = sql.SqlSelectMultiResult(qry);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("getDataSetFromDB on Error : " + ex.Message);
                    rstDs = null;
                }
            }

            return rstDs;
        }

    }



}
