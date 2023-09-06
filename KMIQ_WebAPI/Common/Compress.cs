using System;
using System.Text;
using System.IO;
using System.IO.Compression;

namespace loCommon
{
    public class Compress
    {
        /// <summary>
        /// string을 압축하고 Base64문자열로 리턴한다.
        /// </summary>
        /// <param name="preCompressedText">압축할 문자열</param>
        /// <returns></returns>
        public static string Compression(string preCompressedText)
        {
            var rowData = Encoding.UTF8.GetBytes(preCompressedText);
            byte[] compressed = null;
            using (var outStream = new MemoryStream())
            {
                using (var hgs = new GZipStream(outStream, CompressionMode.Compress))
                {
                    //outStream에 압축을 시킨다.
                    hgs.Write(rowData, 0, rowData.Length);
                }
                compressed = outStream.ToArray();
            }

            return Convert.ToBase64String(compressed);
        }

        /// <summary>
        /// 압축된 Base64문자열의 압축을 해제하고 기본 string으로 리턴한다.
        /// </summary>
        /// <param name="compressedStr"></param>
        /// <returns></returns>
        public static string DeCompression(string compressedStr)
        {
            string output = null;
            byte[] cmpData = Convert.FromBase64String(compressedStr);
            using (var decomStream = new MemoryStream(cmpData))
            {
                using (var hgs = new GZipStream(decomStream, CompressionMode.Decompress))
                {
                    //decomStream에 압축 헤제된 데이타를 저장한다.
                    using (var reader = new StreamReader(hgs))
                    {
                        output = reader.ReadToEnd();
                    }
                }
            }

            return output;
        }


        /// <summary>
        /// Byte array를 압축하고 Base64문자열로 리턴한다.
        /// </summary>
        /// <param name="preCompressedByteArray">압축할 byte array</param>
        /// <returns></returns>
        public static string Compression(byte[] preCompressedByteArray)
        {
            var rowData = preCompressedByteArray;
            byte[] compressed = null;
            using (var outStream = new MemoryStream())
            {
                using (var hgs = new GZipStream(outStream, CompressionMode.Compress))
                {
                    //outStream에 압축을 시킨다.
                    hgs.Write(rowData, 0, rowData.Length);
                }
                compressed = outStream.ToArray();
            }

            return Convert.ToBase64String(compressed);
            //return Encoding.Default.GetString(compressed);
        }


        /// <summary>
        /// 압축된 Base64문자열의 압축을 해제하고 byte array로 리턴한다.
        /// </summary>
        /// <param name="compressedStr">Base64 인코딩 압축문자열</param>
        /// <returns></returns>
        public static byte[] DeCompressionToByteArray(string compressedStr)
        {
            byte[] output = null;
            byte[] cmpData = Convert.FromBase64String(compressedStr);

            using (var decomStream = new MemoryStream(cmpData))
            {
                using (var hgs = new GZipStream(decomStream, CompressionMode.Decompress))
                {
                    hgs.Read(output, 0, (int)hgs.BaseStream.Length);
                    //using (var memstream = new MemoryStream())
                    //{
                    //    hgs.BaseStream.CopyTo(memstream);
                    //    output = memstream.ToArray();
                    //}
                }
            }

            return output;
        }


        /// <summary>
        /// 바이너리 파일을 압축하고 Base64문자열로 리턴한다.
        /// </summary>
        /// <param name="FileName">압축할 파일명(Full path)</param>
        public static string CompressionFile(string FileName)
        {
            string compressed = "";

            if (File.Exists(FileName))
            {
                byte[] fileBytes = null;
                try
                {
                    using (FileStream fs = new FileStream(FileName, FileMode.Open))
                    {
                        fileBytes = new byte[fs.Length];
                        fs.Read(fileBytes, 0, fileBytes.Length);
                    }

                    compressed = Compression(fileBytes);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("CompressionFile on Error : " + ex.Message);
                }
            }

            return compressed;
        }



        /// <summary>
        /// 압축된, 혹은 압축되지 않은 Base64문자열을 특정 파일로 저장시킨다.
        /// </summary>
        /// <param name="compressedStr"></param>
        /// <param name="compressed"></param>
        /// <param name="targetFileName"></param>
        /// <returns></returns>
        public static bool DeCompressionToFile(string compressedStr, bool compressed, string targetFileName)
        {
            bool rtn = false;

            byte[] output = null;
            byte[] cmpData = Convert.FromBase64String(compressedStr);

            if (compressed)
            {
                using (var decomStream = new MemoryStream(cmpData))
                {
                    using (var hgs = new GZipStream(decomStream, CompressionMode.Decompress))
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            hgs.CopyTo(ms);
                            output = ms.ToArray();
                        }
                    }
                }
            }
            else
            {
                output = cmpData;
            }

            if (output.Length > 0)
            {
                try
                {
                    File.WriteAllBytes(targetFileName, output);
                    rtn = true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Write file on Error : " + ex.Message);
                }
            }

            return rtn;
        }

    }
}
