using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GateWay
{
    public class MsgMail
    {
        private Outlook.Application oApp;
        private Outlook._MailItem oMsg;
        private Outlook.Attachment oAttach;
        private Outlook.Recipients oRecips;
        private Outlook.Recipient oRecip;

        public MsgMail()
        {
            oApp = new Outlook.Application();
        }

        public void SendNotice(string mailAddress, string fileName, string unitName, string clientName)
        {
            string[] addressArray = mailAddress.Split(new char[] { ',', ';', '/' });

            if (!(addressArray.Length > 0))
            {
                Global.OutputLine(string.Format("### ошибка! Пустой email адрес: '{0}' для клиента: '{1}({2})'", mailAddress, clientName, unitName));
                Global.IsNoncriticzlError = true;
                return;
            }

            try
            {
                oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                oMsg.Subject = "Извещение 1 для клиента " + clientName + "(" + unitName + ")";
                string strBody = null;
                //foreach (string line in Global.letterData.LetterList)
                //{
                //    strBody += "<h3>" + line + "</h3>";
                //}
                strBody += "<h3>Текст письма </h3>";
                strBody += "<h3>Текст письма </h3>";
                strBody += "<h3>Текст письма </h3>";
                oMsg.HTMLBody = strBody;
                string sDisplayName = "MyAttachment";
                int iPosition = (int)oMsg.Body.Length + 1;
                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                oAttach = oMsg.Attachments.Add(@fileName, iAttachType, iPosition, sDisplayName);
                oRecips = (Outlook.Recipients)oMsg.Recipients;
                foreach (string mail in addressArray)
                {
                    if (!string.IsNullOrWhiteSpace(mail))
                    {
                        oRecip = (Outlook.Recipient)oRecips.Add(mail.Trim());
                    }
                }
                oRecip.Resolve();
                oMsg.Send();
                Global.OutputLine(string.Format("!!! Успешно отправлено Извещение 1 клиенту '{0}({1})'", clientName, unitName));
            }
            catch (Exception ex)
            {
                Global.OutputLine(string.Format("### ошибка при попытке выполнить Mail.SendNotice: '{0}' на адрес: '{1}'", ex.Message, mailAddress));
                Global.IsNoncriticzlError = true;
            }
        }

        public void SendPassport(string mailAddress, string fileName)
        {
            string[] addressArray = mailAddress.Split(new char[] { ',', ';', '/' });

            if (!(addressArray.Length > 0))
            {
                Console.WriteLine(string.Format("*** Ошибка! При выполнении Mail.SendPassport. Пустой email адрес: '{0}'", mailAddress));
                Console.WriteLine(string.Format("Для завершения нажмите Enter"));
                Console.ReadLine();
                return;
            }

            try
            {
                oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                oMsg.Subject = "CO: Паспорт прогона " + Global.WinTitleName;
                string strBody = "<h3>Во вложенном файле отмечены ошибки: фатальные *** и некритичные ###</h3>"; ;
                oMsg.HTMLBody = strBody;
                string sDisplayName = "MyAttachment";
                int iPosition = (int)oMsg.Body.Length + 1;
                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                oAttach = oMsg.Attachments.Add(@fileName, iAttachType, iPosition, sDisplayName);
                oRecips = (Outlook.Recipients)oMsg.Recipients;
                foreach (string mail in addressArray)
                {
                    if (!string.IsNullOrWhiteSpace(mail))
                    {
                        oRecip = (Outlook.Recipient)oRecips.Add(mail.Trim());
                    }
                }
                oRecip.Resolve();
                oMsg.Send();
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("*** Ошибка при попытке выполнить Mail.SendPassport: '{0}' на адрес: '{1}'", ex.Message, mailAddress));
                Console.WriteLine(string.Format("Для завершения нажмите Enter"));
                Console.ReadLine();
            }
        }
    }
}
