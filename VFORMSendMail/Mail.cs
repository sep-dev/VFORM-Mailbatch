using System;
using System.Collections.Generic;
using System.Configuration;
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
using System.Data;
using System.Globalization;

namespace VFORMSendMail
{
    public static class MailTemplate
    {
        public const string SUBJECT_REQUEST_DATA_LIMIT = "【VFORM 保守サービス】申請対象の現物を送付ください";
        public const string SUBJECT_REQUEST_CONTENT_LIMIT = "【VFORM 保守サービス】交換数・校正数の残りがあります";
    }

    public class Mail
    {
        ////送信者名
        //public string NameFrom { get; set; }
        //送信者アドレス
        public string AddressFrom { get; set; }
        //宛先名
        public string NameTo { get; set; }
        //宛先アドレス
        public string AddressTo { get; set; }
        //件名
        public string Subject { get; set; }
        //本文
        public string BodyText { get; set; }

        public List<string> AttachFilePaths { get; set; } = new List<string>();

        // メール送信
        public async Task SendEmailNew()
        {
            string clientId = ConfigurationManager.AppSettings["GmailClientId"].ToString();
            string clientSecret = ConfigurationManager.AppSettings["GmailSecretId"].ToString();
            string refreshToken = ConfigurationManager.AppSettings["GmailRefreshToken"].ToString();
            string AuthUserName = ConfigurationManager.AppSettings["MailAuthUserName"].ToString();

            var token = new TokenResponse
            {
                RefreshToken = refreshToken
            };

            var secrets = new ClientSecrets
            {
                ClientId = clientId,
                ClientSecret = clientSecret
            };

            var credential = new UserCredential(
                new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
                {
                    ClientSecrets = secrets
                }),
                AuthUserName,
                token
            );

            bool refreshed = await credential.RefreshTokenAsync(CancellationToken.None);
            if (!refreshed)
            {
                throw new Exception("トークンのリフレッシュに失敗しました。");
            }

            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Gmail API Sender"
            });

            MimeMessage MimeMessage = new MimeMessage();
            MimeMessage.From.Add(new MailboxAddress(ConfigurationManager.AppSettings["MailFrom"].ToString(), AddressFrom));
            foreach (string Address in AddressTo.Split(','))
            {
                MimeMessage.To.Add(new MailboxAddress("", Address));
            }
            MimeMessage.Subject = Subject;

            TextPart bodyPart;
            bodyPart = new TextPart("plain")
            {
                Text = BodyText,
                ContentTransferEncoding = ContentEncoding.Base64,
            };
            bodyPart.ContentType.Charset = "utf-8";
            MimeMessage.Body = bodyPart;

            var message = new Google.Apis.Gmail.v1.Data.Message
            {
                Raw = Base64UrlEncode(MimeMessage.ToString())
            };

            var result = await service.Users.Messages.Send(message, "me").ExecuteAsync();
        }

        private static string Base64UrlEncode(string input)
        {
            var bytes = Encoding.UTF8.GetBytes(input);
            return Convert.ToBase64String(bytes)
                .Replace("+", "-")
                .Replace("/", "_")
                .Replace("=", "");
        }


        private void WriteErrorLog(Exception ex)
        {
            StringBuilder objBld = new StringBuilder();
            objBld.AppendLine("メールの送信が失敗しました。\t" + DateTime.Now.ToString());
            objBld.AppendLine(ex.Message);
            objBld.AppendLine("");
            objBld.AppendLine("");
            objBld.AppendLine("AddressFrom：" + AddressFrom);
            objBld.AppendLine("AddressTo：" + AddressTo);
            objBld.AppendLine("Subject：" + Subject);
            objBld.AppendLine("BodyText：" + BodyText);

            CommonLog.Instance.WriteLog(objBld.ToString());
        }

        public static class VformMail
        {
            public static void SendMailRequestDateLimit(DataTable RequestDateLimitUsers)
            {
                if (RequestDateLimitUsers.Rows.Count == 0) { return; }

                StringBuilder MailBody = new StringBuilder();
                DateTime CurrentDate = DateTime.Now;
                Mail Mail = new Mail();

                foreach (DataRow Row in RequestDateLimitUsers.Rows)
                {
                    Mail = new Mail();
                    MailBody.Clear();
                    MailBody.AppendLine(Row["user_name"].ToString() + " 様");
                    MailBody.AppendLine("");
                    MailBody.AppendLine("VFORM管理者からお知らせいたします。");
                    MailBody.AppendLine("交換・校正のご申請をいただいておりますが、対象物を送付いただけておりません。");
                    MailBody.AppendLine("期日までのご送付をお願いいたします。");
                    MailBody.AppendLine("※送付タイミングの行き違い連絡となってしまった場合は、ご容赦ください。");
                    MailBody.AppendLine("");
                    MailBody.AppendLine("■ご申請内容");
                    MailBody.AppendLine("送付期日：" + Row["send_limit_date"].ToString());
                    if (! string.IsNullOrEmpty(Row["target"].ToString()))
                    {
                        MailBody.AppendLine("交換ターゲットNo：" + Row["target"].ToString());
                    }
                    if (!string.IsNullOrEmpty(Row["camera"].ToString()))
                    {
                        MailBody.AppendLine("構成カメラNo：" + Row["camera"].ToString());
                    }

                    Mail.AddressFrom = ConfigurationManager.AppSettings["VformSystemMail"].ToString();
                    Mail.AddressTo = Row["mail"].ToString();
                    Mail.Subject = MailTemplate.SUBJECT_REQUEST_DATA_LIMIT;
                    Mail.BodyText = MailBody.ToString();
                    try
                    {
                        // 非同期メソッドを同期的に実行
                        Task.Run(() =>
                            Mail.SendEmailNew()
                        ).GetAwaiter().GetResult();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("エラーが発生しました: " + ex.Message);
                    }
                }
            }
            //暫定
            public static void SendMailContentLimit(DataTable ContentLimitUsers)
            {
                if (ContentLimitUsers.Rows.Count == 0) { return; }

                StringBuilder MailBody = new StringBuilder();
                DateTime CurrentDate = DateTime.Now;
                Mail Mail = new Mail();

                foreach (DataRow Row in ContentLimitUsers.Rows)
                {
                    Mail = new Mail();
                    MailBody.Clear();
                    MailBody.AppendLine("VFORM管理者からお知らせいたします。");
                    MailBody.AppendLine("ご契約されているターゲット交換数・カメラ校正数が残存していることを確認いたしました。");
                    MailBody.AppendLine("契約の期限が近くなっているため、交換・校正の必要がございましたら、ご考慮ください。");
                    MailBody.AppendLine("※送付タイミングの行き違い連絡となってしまった場合は、ご容赦ください。");
                    MailBody.AppendLine("");
                    MailBody.AppendLine("■契約内容");
                    MailBody.AppendLine("契約期限：" + Row["request_limit_date"].ToString());
                    MailBody.AppendLine("契約番号：" + Row["contract_id"].ToString());
                    MailBody.AppendLine("契約名：" + Row["contract_type_name"].ToString());
                    MailBody.AppendLine("口数：" + Row["contract_count"].ToString());
                    if (!string.IsNullOrEmpty(Row["target_count"].ToString()))
                    {
                        if (int.Parse(Row["target_count"].ToString()) > 0)
                        {

                            MailBody.AppendLine("ターゲット交換残数：" + Row["target_count"].ToString());
                        }
                    }
                    if (!string.IsNullOrEmpty(Row["camera_count"].ToString()))
                    {
                        if (int.Parse(Row["camera_count"].ToString()) > 0)
                        {
                            MailBody.AppendLine("カメラ校正残数：" + Row["camera_count"].ToString());
                        }
                    }

                    Mail.AddressFrom = ConfigurationManager.AppSettings["VformSystemMail"].ToString();
                    Mail.AddressTo = Row["mail"].ToString();
                    Mail.Subject = MailTemplate.SUBJECT_REQUEST_CONTENT_LIMIT;
                    Mail.BodyText = MailBody.ToString();
                    try
                    {
                        // 非同期メソッドを同期的に実行
                        Task.Run(() =>
                            Mail.SendEmailNew()
                        ).GetAwaiter().GetResult();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("エラーが発生しました: " + ex.Message);
                    }
                }
            }
        }
    }
}
