using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Configuration;
using System.IO;
using System.Threading;

namespace ExcelSol.Pages
{
    public class EmailPage : TestFixtureBase
    {
        static Outlook.Application application = new Outlook.Application();
        static Outlook.Accounts accounts = application.Session.Accounts;

        public static void searchEmailAndDownlaodAttachments(string subjects, out string fileName, out string url, bool saveAttachement = false)
        {
            fileName = "";
            url = "";
            List<string> linksSylndreList = new List<string>();

            foreach (Outlook.Account account in accounts)
            {
                if (string.Equals(account.DisplayName.Trim(), "t-mamaher@EFG-HERMES.com", StringComparison.InvariantCultureIgnoreCase))
                {
                    Console.WriteLine("Current email account is: " + account.DisplayName);

                    Outlook.Application oApp = new Outlook.Application();
                    Outlook.NameSpace oNS = oApp.GetNamespace("mapi");

                    oNS.SendAndReceive(false);
                    Thread.Sleep(5000);

                    oNS.Logon(Missing.Value, Missing.Value, false, true);

                    Outlook.MAPIFolder oInbox = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    //Outlook.MAPIFolder oCylnderFolder = ((Outlook.MAPIFolder)oInbox).Folders["Cylnder"];
                    Outlook.Items oItems = oInbox.Items;

                    oItems.Sort("[ReceivedTime]", false);

                    Outlook.MailItem oMsg = (Outlook.MailItem)oItems[oItems.Count];

                    //Check for attachments.
                    int AttachCnt = 0;
                    string filename = "";
                    bool emailFound = false;
                    int itr = 0;

                    while (emailFound == false)
                    {
                        for (int i = oItems.Count; i >= 1; i--)
                        {
                            try
                            {
                                oMsg = (Outlook.MailItem)oItems[i];

                                //Output some common properties.
                                string subj = oMsg.Subject;
                                string sender = oMsg.SenderName;
                                string reciTime = oMsg.ReceivedTime.ToString();
                                string Body = oMsg.Body;
                                string HTMLBody = oMsg.HTMLBody;

                                DateTime reciTimeDT = DateTime.Parse(reciTime);

                                //if (reciTimeDT.ToString("MM/dd/yyy") != DateTime.Now.ToString("MM/dd/yyy"))
                                //    break;

                                if (!String.IsNullOrEmpty(subj))
                                {
                                    if (subj == subjects.Trim())
                                    {
                                        //url = linksSylndreList[0];
                                        emailFound = true;
                                        AttachCnt = oMsg.Attachments.Count;
                                        sender = oMsg.SenderName;
                                        reciTime = oMsg.ReceivedTime.ToString();

                                        //if (saveAttachement)
                                        for (int j = AttachCnt; j >= 1; j--)
                                        {
                                            Outlook.Attachment attachment = oMsg.Attachments[j];
                                            
                                            var application_path = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
                                            string applicationPath_new = application_path.Replace("\\bin\\Debug", "");
                                            string applicationPath_new_name = applicationPath_new + "DownLoads";
                                            
                                            filename = Path.Combine(applicationPath_new_name, attachment.FileName);
                                            attachment.SaveAsFile(filename);
                                            Thread.Sleep(5000);
                                        }

                                        // oItems[i].Move(oCylnderFolder);
                                        break;
                                    }
                                }
                            }

                            catch { }
                        }

                        if (emailFound == false)
                        {
                            oNS.SendAndReceive(false);
                            Thread.Sleep(5000);
                        }

                        itr += 1;

                        if (itr == 3)
                            break;
                    }

                    Thread.Sleep(2000);
                    Console.WriteLine("The attchmnet file is downloaded successfully");
                    fileName = filename;
                }
            }
        }

        public static void sendEmailResults(string subject,string to, string cc, string body, bool sendAttachment = false)
        {
            foreach (Account account in accounts)
            {
                if (string.Equals(account.DisplayName.Trim(), "t-mamaher@EFG-HERMES.com", StringComparison.InvariantCultureIgnoreCase))
                {
                    var mail = ((MailItem)application.CreateItem(OlItemType.olMailItem));

                    mail.Subject = subject;
                    mail.To = to;
                    mail.CC = cc;
                    mail.HTMLBody = body;

                    string attachFilePath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase.Replace("\\bin\\Debug", "\\Results") + "New Text Document.txt";

                    if (sendAttachment)
                        mail.Attachments.Add(attachFilePath, OlAttachmentType.olByValue, Type.Missing, Type.Missing);

                    mail.SendUsingAccount = account;
                    mail.Send();
                    Thread.Sleep(5000);
                }
            }
        }
    }
}
