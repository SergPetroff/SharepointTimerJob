using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System.Net.Mail;
using System.Net;
using System.Text.RegularExpressions;

namespace UCG_TJob.TimerJobs
{
    public class InfoFromManager:SPJobDefinition
    {
        public InfoFromManager()
            : base()
            {
            }
        public InfoFromManager(string jobName, SPService service): base(jobName, service,null, SPJobLockType.None)
        {
            this.Title = "Уведомление управляющим";

        }

        public InfoFromManager(string jobName, SPWebApplication webapp)
            : base(jobName, webapp, null, SPJobLockType.ContentDatabase)
        {
            this.Title = "Уведомление управляющим";

        }


        public override void Execute(Guid targetInstanceId)
        {
            // запускаем только в понедельник и воскресенье
            DateTime CurrentDay = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            if (CurrentDay.DayOfWeek == DayOfWeek.Monday || CurrentDay.DayOfWeek == DayOfWeek.Sunday)
            {
                checkReportFromManager();
            }
            
        }

        private void checkReportFromManager()
        {
            
            using (SPSite site = new SPSite("http://xrm"))
            {
                using (SPWeb webApp = site.OpenWeb("crm"))
                {
                    Guid guidObj = new Guid("6318a62c-dfbd-4452-af42-919e3c66f011"); //получаем список обьекты по GUID
                    SPList listObj = webApp.Lists[guidObj];
                    SPListItemCollection spListObjectsColl = listObj.Items;

                    Guid guidInfoList = new Guid("043891f9-e187-4365-9d80-dfa908f5ced0"); //получаем список записей от менеджеров по GUID
                    SPList listInfo = webApp.Lists[guidInfoList];
                    SPListItemCollection spListInfoColl = listInfo.Items;

                    DateTime CurrentDay = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day).AddDays(-7);

                     //получаем записи из списка обьектов 
                    var listObjects = from i in spListObjectsColl.OfType<SPListItem>()
                                      where i["Status"].ToString() == "Текущий" && i["Manager"] != null
                                      select i;

                    //записи из списка с информацией от управляющих созданых за последние 7 дней.
                    var listqueryInfoColl = from i in spListInfoColl.OfType<SPListItem>()
                                            where ((DateTime)i["Created"]).Date >= CurrentDay && i["Object"] != null
                                            select i;
                     //создаем список из ИД оьтектов списка информации от управляющих
                    List<int> listidobjects = new List<int>();
                    foreach (SPListItem item in listqueryInfoColl)
                    {
                        SPFieldLookupValue fieldLookupValue = new SPFieldLookupValue(item["Object"].ToString());
                        int lookupIDObject = fieldLookupValue.LookupId;
                        listidobjects.Add((int)lookupIDObject);

                    }


                    // выбираем обьекты у которых нет записей
                    var noexistObjectList = from itemObj in listObjects
                                            where !listidobjects.Any(infoitem => (infoitem.ToString() == itemObj["ID"].ToString()))
                                            select itemObj;

                     //перебираем записи, от правляем уведомления управляющим
                    foreach (SPListItem item in noexistObjectList)
                    {
                        SPFieldUserValue managercantin = new SPFieldUserValue(webApp, item["Manager"].ToString()); //учетная запись управляющего

                        if (IsValidEmailId(managercantin.User.Email))
                        {

                            sendEmail(managercantin.User.Email, "Журнал - Информация от управляющих", managercantin.User.Name);
                        }
                    }

                }
            }
        }

        public static bool IsValidEmailId(string InputEmail)
        {
            Regex regex = new Regex(@"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,3})+)$");
            Match match = regex.Match(InputEmail);
            if (match.Success)
                return true;
            else
                return false;
        }


        private void sendEmail(string Mailaddres, string subject, string body)
        {
            using (MailMessage mail = new MailMessage())
            {
                string bodyEmailHeader = "";
                string bodyEmailFooter = "";
                bodyEmailHeader = "<HTML><HEAD><META name=GENERATOR content='MSHTML 11.00.9600.17344'></HEAD> <BODY style='FONT-SIZE: 11pt; FONT-FAMILY: Segoe UI Light,sans-serif; COLOR: #444444'>";
                bodyEmailHeader += "<DIV><SPAN style='FONT-SIZE: 13.5pt'>Уважаемый " + body + " Вы не вносили запись в журнал  'Информация от управляющих' уже более 7 дней</SPAN></DIV><DIV>&nbsp;</DIV><table><tbody>";
                bodyEmailFooter = "</table></tbody><P><A href='http://xrm/CRM/Lists/InfofromManagers'>Перейти к журналу</A></P></BODY></HTML>";

                mail.From = new MailAddress("UCG<portal@ucg.ru>");
                mail.To.Add(Mailaddres);
                mail.Subject = subject;
                mail.IsBodyHtml = true;
                mail.Body = bodyEmailHeader + bodyEmailFooter;

                SmtpClient smtp = new SmtpClient("smtp.ucg.ru");

                smtp.UseDefaultCredentials = true;
                //smtp.UseDefaultCredentials = false;
                //smtp.Credentials = new NetworkCredential("portal@ucg.ru", "mega_2013_av");
                smtp.Send(mail);
            }
        }

    }
}
