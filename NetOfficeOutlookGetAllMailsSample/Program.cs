using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using NetOffice.OutlookApi;
using System.Linq.Expressions;

namespace Fu.NetOfficeOutlookGetAllMailsSample
{
    internal class Program
    {
        /// <summary>
        /// 保存フォルダー
        /// </summary>
        private static string _folder = ROOT_FOLDER;
        private const string ROOT_FOLDER = "受信フォルダー";

        /// <summary>
        /// Entry Point
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            // Outlookインスタンス作成
            var outlookApplication = new Outlook.Application();

            // MAPI取得
            // 環境によっては事前にOutlook起動が必要？
            var outlookNS = outlookApplication.GetNamespace("MAPI");

            // 存在するアカウントで繰り返し
            foreach (var account in outlookNS.Session.Accounts)
            {
                // アカウント取得
                var acc = account as Outlook.Account;

                // アカウントの受信フォルダー取得
                var inboxFolder = account.DeliveryStore.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                // メール数取得
                System.Console.WriteLine($"アカウント[{acc.DisplayName}]:フォルダー内メール数:{inboxFolder.Items.Count}");

                // メール内容取得
                GetMails(inboxFolder);

                // 出力用のフォルダー階層クリア
                _folder = ROOT_FOLDER;

                // Dispose
                inboxFolder.Dispose();
                acc.Dispose();
                account.Dispose();
            }

            // 破棄
            outlookNS.Session.Accounts.Dispose();
            outlookNS.Session.Dispose();
            outlookNS.Dispose();

            // 破棄
            outlookApplication.Quit();
            outlookApplication.Dispose();
        }

        /// <summary>
        /// メール取得
        /// </summary>
        /// <param name="folder"></param>
        static void GetMails(MAPIFolder folder)
        {
            // フォルダー内のメール情報出力
            foreach (var item in folder.Items)
            {
                // 通常のメールと会議をとりあえずサポート
                var mailitem = item as Outlook.MailItem;
                var meetingitem =  item as MeetingItem;

                // メールの場合
                if (mailitem != null)
                {
                    System.Console.WriteLine($"受信時刻[{mailitem.ReceivedTime}]:題名[{mailitem.Subject}]");
                    foreach (var attachments in mailitem.Attachments)
                    {
                        System.Console.WriteLine($"添付ファイル:{attachments.DisplayName}");
                        attachments.Dispose();
                    }
                    // 破棄
                    mailitem.Attachments.Dispose();
                    mailitem.Dispose();
                }

                // 会議の場合
                if (meetingitem != null)
                {
                    System.Console.WriteLine($"受信時刻[{meetingitem.ReceivedTime}]:題名[{meetingitem.Subject}]:本文[{meetingitem.Body}]");

                    // 添付ファイル取得
                    foreach (var attachments in meetingitem.Attachments)
                    {
                        System.Console.WriteLine($"添付ファイル:{attachments.DisplayName}");
                        attachments.Dispose();
                    }
                    // 破棄
                    meetingitem.Attachments.Dispose();
                    meetingitem.Dispose();
                }
            }

            // 下位のフォルダー取得→再帰実行
            foreach (var nowDirFolder in folder.Folders)
            {
                var folderitem = nowDirFolder as MAPIFolder;

                // 今いる階層と今回のフォルダーを連結
                _folder += "/" + folderitem.Name;
                System.Console.WriteLine($"フォルダー[{_folder}]メール数:{folderitem.Items.Count}");

                // 再帰
                GetMails(folderitem as MAPIFolder);

                // 破棄
                folderitem.Dispose();
            }

            // 破棄
            folder.Dispose();

            // フォルダーがなければクリア
            _folder = ROOT_FOLDER;
        }
    }
}
