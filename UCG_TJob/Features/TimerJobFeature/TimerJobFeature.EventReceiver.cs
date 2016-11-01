using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Administration;

namespace UCG_TJob.Features.TimerJobFeature
{
    /// <summary>
    /// Этот класс обрабатывает события, возникающие в ходе активации, деактивации, установки, удаления и обновления компонентов.
    /// </summary>
    /// <remarks>
    /// GUID, присоединенный к этому классу, может использоваться при создании пакета и не должен изменяться.
    /// </remarks>

    [Guid("2f9a1e9f-cfd7-469a-97fb-2c86c8efcffe")]
    public class TimerJobFeatureEventReceiver : SPFeatureReceiver
    {
        // Раскомментируйте ниже метод для обработки события, возникающего после активации компонента.
        const string JobName = "Уведомление управляющим";
        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    SPWebApplication parentWebApp = (SPWebApplication)properties.Feature.Parent;
                    DeleteExistingJob(JobName, parentWebApp);
                    CreateJob(parentWebApp);
                });  
            }
            catch (Exception ex) 
            {
                throw ex; 
            }
        }

        private bool CreateJob(SPWebApplication site)
        {
            bool jobCreated = false;
            try
            {
                TimerJobs.InfoFromManager job = new TimerJobs.InfoFromManager(JobName, site);
                //SPWeeklySchedule schedule = new SPWeeklySchedule();
                //schedule.BeginDayOfWeek = DayOfWeek.Saturday;
                //schedule.BeginHour = 8;
                //schedule.BeginMinute = 0;
                //schedule.BeginSecond = 0;
                //schedule.EndSecond = 0;
                //schedule.EndMinute = 10;
                //schedule.EndHour = 8;
                //schedule.EndDayOfWeek = DayOfWeek.Saturday;
                //job.Schedule = schedule;
                //job.Update();

                //SPMinuteSchedule schedule = new SPMinuteSchedule();
                //schedule.BeginSecond = 0;
                //schedule.EndSecond = 59;
                //schedule.Interval = 1;
                //job.Schedule = schedule;


                SPDailySchedule schedule = new SPDailySchedule();
                schedule.BeginHour = 8;
                schedule.BeginMinute = 0;
                schedule.BeginSecond = 0;

                schedule.EndHour = 8;
                schedule.EndMinute = 15;
                schedule.EndSecond = 00;

                job.Schedule = schedule;
                job.Update();
            }
            catch (Exception)
            {
                return jobCreated;
            }
            return jobCreated;
        }
        public bool DeleteExistingJob(string jobName, SPWebApplication site)
        {
            bool jobDeleted = false;
            try
            {
                foreach (SPJobDefinition job in site.JobDefinitions)
                {
                    if (job.Name == jobName)
                    {
                        job.Delete();
                        jobDeleted = true;
                    }
                }
            }
            catch (Exception)
            {
                return jobDeleted;
            }
            return jobDeleted;
        }

        // Раскомментируйте ниже метод для обработки события, возникающего перед деактивацией компонента.

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {

            lock (this)
            {
                try
                {
                    SPSecurity.RunWithElevatedPrivileges(delegate()
                    {
                        SPWebApplication parentWebApp = (SPWebApplication)properties.Feature.Parent;
                        DeleteExistingJob(JobName, parentWebApp);
                    });
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }


        // Раскомментируйте ниже метод для обработки события, возникающего после установки компонента.

        //public override void FeatureInstalled(SPFeatureReceiverProperties properties)
        //{
        //}


        // Раскомментируйте ниже метод для обработки события, возникающего перед удалением компонента.

        //public override void FeatureUninstalling(SPFeatureReceiverProperties properties)
        //{
        //}

        // Раскомментируйте ниже метод для обработки события, возникающего при обновлении компонента.

        //public override void FeatureUpgrading(SPFeatureReceiverProperties properties, string upgradeActionName, System.Collections.Generic.IDictionary<string, string> parameters)
        //{
        //}
    }
}
