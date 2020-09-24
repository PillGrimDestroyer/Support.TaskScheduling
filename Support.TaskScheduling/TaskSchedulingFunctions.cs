using Microsoft.Win32.TaskScheduler;
using System;
using System.IO;
using System.Security.Principal;

namespace Support.TaskScheduling
{
    public class TaskSchedulingFunctions
    {
        public static bool Disable(string TaskName)
        {
            if (!IsAdministrator())
                return false;

            using (TaskService taskService = new TaskService())
            {
                Task task = taskService.FindTask(TaskName);

                if (task == null)
                {
                    return false;
                }

                task.Enabled = false;
                return true;
            }
        }

        public static bool Delete(string TaskName)
        {
            if (!IsAdministrator())
                return false;

            using (TaskService taskService = new TaskService())
            {
                Task task = taskService.FindTask(TaskName);

                if (task == null)
                {
                    return false;
                }

                taskService.RootFolder.DeleteTask(TaskName);
                return true;
            }
        }

        public static bool CreateOrActivate()
        {
            string CurrentProcessFullFileName = System.Diagnostics.Process.GetCurrentProcess()
                    .MainModule.FileName.Replace("vshost.", "");

            return CreateOrActivate(Path.GetFileNameWithoutExtension(CurrentProcessFullFileName),
                CurrentProcessFullFileName, TimeSpan.FromHours(1));
        }

        public static bool CreateOrActivate(string TaskName, int RetryAllMinutes)
        {
            return CreateOrActivate(TaskName, System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName,
                TimeSpan.FromMinutes(RetryAllMinutes));
        }

        public static bool CreateOrActivate(string TaskName, string FullExecuteableFileName, TimeSpan RetryIntervall)
        {
            using (TaskService taskService = new TaskService())
            {
                Task task = taskService.FindTask(TaskName);

                if (task == null)
                {
                    TaskDefinition taskDefinition = CreateTask(taskService, TaskName, RetryIntervall);

                    taskDefinition.RegistrationInfo.Description = $"Support created Task for " + $"{Path.GetFileNameWithoutExtension(FullExecuteableFileName)}";
                    taskDefinition.Settings.MultipleInstances = TaskInstancesPolicy.IgnoreNew;
                    taskDefinition.Settings.AllowDemandStart = true;
                    taskDefinition.Settings.AllowHardTerminate = false;
                    taskDefinition.Settings.DisallowStartIfOnBatteries = false;
                    taskDefinition.RegistrationInfo.Author = "Support";

                    taskDefinition.Actions.Add(new ExecAction(FullExecuteableFileName, workingDirectory: Path.GetDirectoryName(FullExecuteableFileName)));

                    task = taskService.RootFolder.RegisterTaskDefinition(TaskName, taskDefinition);
                }

                task.Enabled = true;
            }

            return true;
        }

        private static TaskDefinition CreateTask(TaskService taskService, string taskName, TimeSpan RetryIntervall)
        {
            DailyTrigger dailyTrigger = new DailyTrigger();
            LogonTrigger logonTrigger = new LogonTrigger();

            dailyTrigger.StartBoundary = DateTime.Now.Date;
            dailyTrigger.Repetition.Interval = RetryIntervall;
            dailyTrigger.Enabled = true;

            //dailyTrigger.EndBoundary = DateTime.MaxValue;
            //dailyTrigger.ExecutionTimeLimit = TimeSpan.MaxValue;

            TaskDefinition taskDefinition = taskService.NewTask();

            taskDefinition.Triggers.Add(dailyTrigger);
            taskDefinition.Triggers.Add(logonTrigger);

            taskDefinition.Principal.RunLevel = TaskRunLevel.Highest;
            taskDefinition.Principal.LogonType = TaskLogonType.InteractiveToken;
            //taskDefinition.Principal.UserId = "Support\\TaskSchedulingFunctions";

            return taskDefinition;
        }

        public static bool IsAdministrator()
        {
            return (new WindowsPrincipal(WindowsIdentity.GetCurrent()))
                    .IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}
