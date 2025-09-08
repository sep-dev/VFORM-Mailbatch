using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using MimeKit;
using System.Configuration;

namespace VFORMSendMail
{
    internal class Program
    {
        static void Main(string[] args)
        {
            DataTable RequestDateLimitUsers = new DataTable();
            DataTable ContentLimitUsers = new DataTable();

            RequestDateLimitUsers = SelectSQL.GetRequestDateLimitUser();
            ContentLimitUsers = SelectSQL.GetContentLimitUser();
            Mail.VformMail.SendMailRequestDateLimit(RequestDateLimitUsers);
            Mail.VformMail.SendMailContentLimit(ContentLimitUsers);
        }
    }


    public class CommonLog
    {
        // マルチスレッド対応のシングルトンパターンを採用
        private static volatile CommonLog instance;
        private static object syncRoot = new Object();
        private CommonLog() { }

        // ロック用のインスタンス
        private static ReaderWriterLock rwl = new ReaderWriterLock();

        // シングルトンパターンのインスタンス取得メソッド
        public static CommonLog Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                            instance = new CommonLog();
                    }
                }
                return instance;
            }
        }

        public void WriteLog(String Log)
        {
            if (Directory.Exists(ConfigurationManager.AppSettings["ErrorLog"].ToString()) == false)
            {
                Directory.CreateDirectory(ConfigurationManager.AppSettings["ErrorLog"].ToString());
            }

            String LogUrl = ConfigurationManager.AppSettings["ErrorLog"].ToString();

            // ログファイルの排他処理
            rwl.AcquireWriterLock(Timeout.Infinite);
            try
            {
                using (StreamWriter streamWriter = new StreamWriter(LogUrl, true))
                {
                    streamWriter.WriteLine(Log);
                }
            }
            finally
            {
                // ロック解除は finally の中で行う
                rwl.ReleaseWriterLock();
            }
        }
    }
}
