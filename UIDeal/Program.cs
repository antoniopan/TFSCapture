using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UIDeal
{
    class Program
    {
        static void Main(string[] args)
        {
            var tfsTracker = new TFS_TRACKER.TfsTracker()
            {
                UserName = "liangliang.pan",
                Password = "2Antonio",
                FileName = args[0],
                FileNameModule = "haha"
            };

            string sQueryTaskExpired = @"select [System.Id], [System.Title], [Microsoft.VSTS.Common.Priority], [System.AssignedTo], [System.State], [MR.Requirement.ApplicativeProject], [System.NodeName], [UI.bug.keyword], [UI.Reserved], [UI.Bug.ExpectFixedBranch], [UI.ExpectedSolvedDate], [Microsoft.VSTS.Common.ResolvedDate], [Microsoft.VSTS.Common.ResolvedBy], [System.CreatedDate], [System.ChangedDate] from WorkItems where [System.TeamProject] = 'Task' and (([System.AreaPath] under 'Task\SW\UIDeal_AWS' and [MR.Requirement.ApplicativeProject] contains '[HSW-uWS-CT]') or [System.AreaPath] under 'Task\SW\uInnovation' or [System.AreaPath] under 'Task\SW\MCSF1' or [System.NodeName] = 'Report') and [System.WorkItemType] = 'Task' and [Microsoft.VSTS.Common.Activity] = 'Improvement' and [UI.Bug.ExpectFixedBranch] contains '[SWBU_69_SP4]' and [System.NodeName] in ('Heart', 'Liver Evaluation', 'Lung Nodule', 'Dual Energy Analysis', 'Dental Application', 'Colon Analysis', 'Body Perfusion', 'Bone Structure Analysis', 'Lung Density Analysis', 'Report', 'Review 4D', 'Dynamic Analysis_CT', 'Vessel Analysis_CT', 'Vessel Analysis_MR', 'Heart_Common', 'Vessel_Common', 'Vessel_Heart_Combined', 'FFR', 'Brain Perfusion_CT_3D') and [System.State] in ('Assigned', 'New') and [UI.ExpectedSolvedDate] < @today order by [System.State], [System.NodeName]".Replace("@today", String.Format("@today-{0:D}", TFS_TRACKER.TfsTracker.FindLastMonday()));
            string sQueryTaskUnReviewed = @"select [System.Id], [System.WorkItemType], [System.Title], [Microsoft.VSTS.Common.Priority], [System.AssignedTo], [System.State], [UI.bug.keyword], [System.NodeName], [UI.ExpectedSolvedDate], [System.CreatedDate] from WorkItems where [System.TeamProject] = 'Task' and (([System.AreaPath] under 'Task\SW\UIDeal_AWS' and [MR.Requirement.ApplicativeProject] contains '[HSW-uWS-CT]') or [System.AreaPath] under 'Task\SW\uInnovation' or [System.AreaPath] under 'Task\SW\MCSF1' or [System.NodeName] = 'Report') and [System.WorkItemType] = 'Task' and [Microsoft.VSTS.Common.Activity] = 'Improvement' and ([System.State] = 'New' or [System.State] = 'Assigned') and [System.NodeName] in ('Heart', 'Liver Evaluation', 'Lung Nodule', 'Dual Energy Analysis', 'Dental Application', 'Colon Analysis', 'Body Perfusion', 'Bone Structure Analysis', 'Lung Density Analysis', 'Report', 'Review 4D', 'Dynamic Analysis_CT', 'Vessel Analysis_CT', 'Vessel Analysis_MR', 'Heart_Common', 'Vessel_Common', 'Vessel_Heart_Combined', 'FFR', 'Brain Perfusion_CT_3D') and not [UI.Bug.ExpectFixedBranch] contains '[SWBU_69_SP4]' and not [UI.Bug.ExpectFixedBranch] contains '[SWBU_71]' order by [System.State], [System.NodeName]";
            //string sQueryExpiredKey = "";

            tfsTracker.InitializeTFS();

            tfsTracker.ExtractTaskList(sQueryTaskExpired);
            tfsTracker.WriteExcelItemList("Improvement Task Expired");

            tfsTracker.ExtractTaskList(sQueryTaskUnReviewed);
            tfsTracker.WriteExcelItemList("Improvement Task Unreviewed");

            //tfsTracker.ExtractTaskList(sQueryExpiredKey);
            //tfsTracker.WriteExcelItemList("Designed Task Expired.");

            tfsTracker.WriteName2File(args[1]);
        }
    }
}
