﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Xml;
namespace TestTFS
{
    static class Program
    {
        static void Main(string[] args)
        {
            var tfsTracker = new TfsTracker()
            {
                UserName = "liangliang.pan",
                Password = "2Antonio",
                FileName = args[0],
                FileNameModule = args[1]
            };

            string sQueryCMTCUR = @"SELECT [System.Id], [System.WorkItemType], [System.NodeName], [XR.UModule], [System.Title], [System.AssignedTo], [System.State], [System.CreatedDate], [Microsoft.VSTS.Scheduling.FinishDate], [Microsoft.VSTS.Scheduling.StartDate], [Microsoft.VSTS.Common.StackRank], [System.IterationPath], [UI.Module], [XR.UAttribute], [XR.Requirement.PanGuSSFS], [XR.UStatus], [XR.URecords], [XR.URemark], [System.AreaPath], [System.CreatedBy], [System.ChangedDate], [Microsoft.VSTS.Common.ResolvedBy], [Microsoft.VSTS.Common.ResolvedDate] FROM WorkItems WHERE [System.TeamProject] = 'XR_Loutang'  AND  [System.WorkItemType] = 'User Requirement'  AND  [System.AreaPath] UNDER 'XR_LouTang\User Requirement\PRODM\SW'  AND  [XR.UAttribute] = '1-CMTC' ORDER BY [XR.UModule] desc";
            string sQueryClinicalUR = @"SELECT [System.Id], [System.WorkItemType], [System.NodeName], [XR.UModule], [System.Title], [System.AssignedTo], [System.State], [System.CreatedDate], [Microsoft.VSTS.Scheduling.FinishDate], [Microsoft.VSTS.Scheduling.StartDate], [Microsoft.VSTS.Common.StackRank], [System.IterationPath], [UI.Module], [XR.UAttribute], [XR.Requirement.PanGuSSFS], [XR.UStatus], [XR.URecords], [XR.URemark], [System.AreaPath], [System.CreatedBy], [System.ChangedDate], [Microsoft.VSTS.Common.ResolvedBy], [Microsoft.VSTS.Common.ResolvedDate] FROM WorkItems WHERE [System.TeamProject] = 'XR_Loutang'  AND  [System.WorkItemType] = 'User Requirement'  AND  [System.AreaPath] UNDER 'XR_LouTang\User Requirement\PRODM\SW'  AND  [XR.UAttribute] IN ('2-临床') ORDER BY [XR.UModule] desc";
            string sQueryAllUR = @"select [System.Id], [System.WorkItemType], [System.NodeName], [XR.UModule], [System.Title], [System.AssignedTo], [System.State], [System.CreatedDate], [Microsoft.VSTS.Scheduling.FinishDate], [Microsoft.VSTS.Scheduling.StartDate], [Microsoft.VSTS.Common.StackRank], [System.IterationPath], [UI.Module], [XR.UAttribute], [XR.Requirement.PanGuSSFS], [XR.UStatus], [XR.URecords], [XR.URemark], [System.AreaPath], [System.CreatedBy], [System.ChangedDate], [Microsoft.VSTS.Common.ResolvedBy], [Microsoft.VSTS.Common.ResolvedDate] from WorkItems where [System.TeamProject] = 'XR_Loutang' and [System.WorkItemType] = 'User Requirement' and [System.AreaPath] under 'XR_LouTang\User Requirement\PRODM\SW' and [XR.UAttribute] in ('1-CMTC', '2-临床') order by [XR.UModule] desc";
            string sQueryAllTask = @"select [System.Id], [System.WorkItemType], [System.Title], [Microsoft.VSTS.Common.Priority], [System.AssignedTo], [System.State], [UI.ExpectedSolvedDate], [System.NodeName], [System.AreaPath], [UI.bug.keyword], [UI.Reserved], [System.CreatedBy], [System.CreatedDate] from WorkItems where [System.TeamProject] = 'Task' and [System.WorkItemType] = 'Task' and [System.AreaPath] under 'Task\XR\08_Loutang\Software' and [Microsoft.VSTS.Common.Activity] = 'Improvement' order by [System.State], [System.AssignedTo], [Microsoft.VSTS.Common.Priority]";
            //string sQueryExpiredUR = @"SELECT [System.Id], [System.WorkItemType], [System.NodeName], [XR.UModule], [System.Title], [System.AssignedTo], [System.State], [System.CreatedDate], [Microsoft.VSTS.Scheduling.FinishDate], [Microsoft.VSTS.Scheduling.StartDate], [Microsoft.VSTS.Common.StackRank], [System.IterationPath], [UI.Module], [XR.UAttribute], [XR.Requirement.PanGuSSFS], [XR.UStatus], [XR.URecords], [XR.URemark], [System.AreaPath], [System.CreatedBy], [System.ChangedDate], [Microsoft.VSTS.Common.ResolvedBy], [Microsoft.VSTS.Common.ResolvedDate] FROM WorkItems WHERE [System.TeamProject] = 'XR_Loutang'  AND  [System.WorkItemType] = 'User Requirement'  AND  [System.AreaPath] UNDER 'XR_LouTang\User Requirement\PRODM\SW'  AND  [XR.UAttribute] = '1-CMTC'  AND  [Microsoft.VSTS.Scheduling.FinishDate] < @today  AND  [System.State] IN ('10-Requirement', '20-Solution', '30-Development')  AND  [Microsoft.VSTS.Scheduling.FinishDate] <> '' ORDER BY [XR.UModule] desc";
            //string sQueryUnPlannedUR = @"SELECT [System.Id], [System.WorkItemType], [System.NodeName], [XR.UModule], [System.Title], [System.AssignedTo], [System.State], [System.CreatedDate], [Microsoft.VSTS.Scheduling.FinishDate], [Microsoft.VSTS.Scheduling.StartDate], [Microsoft.VSTS.Common.StackRank], [System.IterationPath], [UI.Module], [XR.UAttribute], [XR.Requirement.PanGuSSFS], [XR.UStatus], [XR.URecords], [XR.URemark], [System.AreaPath], [System.CreatedBy], [System.ChangedDate], [Microsoft.VSTS.Common.ResolvedBy], [Microsoft.VSTS.Common.ResolvedDate] FROM WorkItems WHERE [System.TeamProject] = 'XR_Loutang'  AND  [System.WorkItemType] = 'User Requirement'  AND  [System.AreaPath] UNDER 'XR_LouTang\User Requirement\PRODM\SW'  AND  [XR.UAttribute] = '1-CMTC'  AND  [System.State] IN ('10-Requirement', '20-Solution', '30-Development')  AND  [Microsoft.VSTS.Scheduling.FinishDate] = '' ORDER BY [XR.UModule] desc";
            //string sQueryExpiredTask = @"SELECT [System.Id], [System.WorkItemType], [System.Title], [Microsoft.VSTS.Common.Priority], [System.AssignedTo], [System.State], [UI.ExpectedSolvedDate], [System.NodeName], [System.AreaPath], [UI.bug.keyword], [UI.Reserved], [System.CreatedBy], [System.CreatedDate] FROM WorkItems WHERE [System.TeamProject] = 'Task'  AND  [System.WorkItemType] = 'Task'  AND  [System.AreaPath] UNDER 'Task\XR\08_Loutang\Software'  AND  [System.NodeName] NOT IN ('SC', 'IC')  AND  [Microsoft.VSTS.Common.Activity] = 'Improvement'  AND  [System.State] IN ('New', 'Assigned')  AND ( [UI.Reserved] CONTAINS 'v0.2'  OR  [UI.Reserved] CONTAINS 'v0.3'  OR  [UI.Reserved] CONTAINS 'v1.0' ) AND  [UI.ExpectedSolvedDate] < @today ORDER BY [UI.ExpectedSolvedDate], [System.NodeName] ";
            //string sQueryUnPlannedTask = @"";
            //string sQueryUnReviewedTask = @"select [System.Id], [System.WorkItemType], [System.Title], [Microsoft.VSTS.Common.Priority], [System.AssignedTo], [System.State], [UI.ExpectedSolvedDate], [System.NodeName], [System.AreaPath], [UI.bug.keyword], [UI.Reserved], [System.CreatedDate], [System.CreatedBy] from WorkItems where [System.TeamProject] = 'Task' and [System.WorkItemType] = 'Task' and [System.AreaPath] under 'Task\XR\08_Loutang\Software' and [Microsoft.VSTS.Common.Activity] = 'Improvement' and ([System.State] = 'Assigned' or [System.State] = 'New') and not [UI.Reserved] contains 'v0.2' and not [UI.Reserved] contains 'v0.3' and not [UI.Reserved] contains 'v1.0' and not [UI.Reserved] contains 'NoCMTC' order by [System.AssignedTo], [System.CreatedDate] desc";
            string sQueryURThisWeek = @"SELECT [System.Id], [System.WorkItemType], [System.NodeName], [XR.UModule], [System.Title], [System.AssignedTo], [System.State], [System.CreatedDate], [Microsoft.VSTS.Scheduling.FinishDate], [Microsoft.VSTS.Scheduling.StartDate], [Microsoft.VSTS.Common.StackRank], [System.IterationPath], [UI.Module], [XR.UAttribute], [XR.Requirement.PanGuSSFS], [XR.UStatus], [XR.URecords], [XR.URemark], [System.AreaPath], [System.CreatedBy], [System.ChangedDate], [Microsoft.VSTS.Common.ResolvedBy], [Microsoft.VSTS.Common.ResolvedDate] FROM WorkItems WHERE [System.TeamProject] = 'XR_Loutang'  AND  [System.WorkItemType] = 'User Requirement'  AND  [System.AreaPath] UNDER 'XR_LouTang\User Requirement\PRODM\SW'  AND  [System.State] IN ('35-Resolved', '40-SSIT Done', '50-SI', '60-SIT', '70-ST')  AND  [Microsoft.VSTS.Common.ResolvedDate] >= @today ORDER BY [XR.UModule] desc".Replace("@today", String.Format("@today-{0:D}", FindLastMonday()));
            string sQueryTaskResolvedThisWeek = @"SELECT [System.Id], [System.WorkItemType], [System.Title], [Microsoft.VSTS.Common.Priority], [System.AssignedTo], [System.State], [UI.ExpectedSolvedDate], [System.NodeName], [System.AreaPath], [UI.bug.keyword], [UI.Reserved], [System.CreatedBy], [System.CreatedDate] FROM WorkItems WHERE [System.TeamProject] = 'Task'  AND  [System.WorkItemType] = 'Task'  AND  [System.AreaPath] UNDER 'Task\XR\08_Loutang\Software'  AND  [Microsoft.VSTS.Common.Activity] = 'Improvement'  AND  [System.State] IN ('Resolved', 'Verified', 'Closed')  AND  [Microsoft.VSTS.Common.ResolvedDate] >= @today ORDER BY [System.State], [System.AssignedTo], [Microsoft.VSTS.Common.Priority]".Replace("@today", String.Format("@today-{0:D}", FindLastMonday()));
            string sQueryTaskCreatedThisWeek = @"SELECT [System.Id], [System.WorkItemType], [System.Title], [Microsoft.VSTS.Common.Priority], [System.AssignedTo], [System.State], [UI.ExpectedSolvedDate], [System.NodeName], [System.AreaPath], [UI.bug.keyword], [UI.Reserved], [System.CreatedBy], [System.CreatedDate] FROM WorkItems WHERE [System.TeamProject] = 'Task'  AND  [System.WorkItemType] = 'Task'  AND  [System.AreaPath] UNDER 'Task\XR\08_Loutang\Software'  AND  [Microsoft.VSTS.Common.Activity] = 'Improvement'  AND  [System.State] IN ('Resolved', 'Verified', 'Closed')  AND  [System.CreatedDate] >= @today ORDER BY [System.State], [System.AssignedTo], [Microsoft.VSTS.Common.Priority]".Replace("@today", String.Format("@today-{0:D}", FindLastMonday()));
            string sQueryOpenP1Task = @"select [System.Id], [System.WorkItemType], [System.Title], [Microsoft.VSTS.Common.Priority], [System.AssignedTo], [System.State], [UI.ExpectedSolvedDate], [System.NodeName], [System.AreaPath], [UI.bug.keyword], [UI.Reserved], [System.CreatedBy], [System.CreatedDate] from WorkItems where [System.TeamProject] = 'Task' and [System.WorkItemType] = 'Task' and [System.AreaPath] under 'Task\XR\08_Loutang\Software' and [Microsoft.VSTS.Common.Activity] = 'Improvement' and [UI.Reserved] contains 'p1' and [System.State] in ('Assigned', 'New') and not [UI.Bug.SolvedBranch] contains '[DSA_CMTC]' order by [System.State], [System.NodeName], [System.AssignedTo] desc";

            tfsTracker.InitializeTFS();

            tfsTracker.ExtractTaskList(sQueryOpenP1Task);
            tfsTracker.WriteExcelItemList("P1 Task List");

            tfsTracker.ExtractTaskQueryInfo(sQueryAllTask);
            tfsTracker.WriteTask2Excel("Task Table");

            tfsTracker.ExtractResolveInfo(sQueryTaskResolvedThisWeek);
            tfsTracker.WriteResolveItemList("Task This Week");

            tfsTracker.ExtractTaskList(sQueryTaskCreatedThisWeek);
            tfsTracker.WriteExcelItemList("Task Created This Week");

            tfsTracker.ExtractURQueryInfo(sQueryCMTCUR);
            tfsTracker.WriteUR2Excel("UR CMTC Table");

            tfsTracker.ExtractURQueryInfo(sQueryClinicalUR);
            tfsTracker.WriteUR2Excel("UR Clinical Table");

            tfsTracker.ExtractURQueryInfo(sQueryAllUR);
            tfsTracker.WriteUrByModule();

            tfsTracker.ExtractResolveInfo(sQueryURThisWeek);
            tfsTracker.WriteResolveItemList("UR This Week");

            tfsTracker.WriteUrByModule();

            tfsTracker.WriteName2File(@"E:\Tracker\name.txt");
        }

        static int FindLastMonday()
        {
            DateTime date = DateTime.Now;
            switch (date.DayOfWeek)
            {
                case System.DayOfWeek.Monday:
                    return 0;
                case System.DayOfWeek.Tuesday:
                    return 1;
                case System.DayOfWeek.Wednesday:
                    return 2;
                case System.DayOfWeek.Thursday:
                    return 3;
                case System.DayOfWeek.Friday:
                    return 4;
                case System.DayOfWeek.Saturday:
                    return 5;
                case System.DayOfWeek.Sunday:
                    return 6;
                default:
                    return 0;
            }
        }
    }
}
