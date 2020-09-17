﻿using System;
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

            string sQueryTaskExpired = @"SELECT [System.Id], [System.Title], [Microsoft.VSTS.Common.Priority], [System.AssignedTo], [System.State], [MR.Requirement.ApplicativeProject], [System.NodeName], [UI.bug.keyword], [UI.Reserved], [UI.Bug.ExpectFixedBranch], [UI.ExpectedSolvedDate], [Microsoft.VSTS.Common.ResolvedDate], [Microsoft.VSTS.Common.ResolvedBy], [System.CreatedDate], [System.ChangedDate] FROM WorkItems WHERE [System.TeamProject] = 'Task'  AND (( [System.AreaPath] UNDER 'Task\SW\UIDeal_AWS'  AND  [MR.Requirement.ApplicativeProject] CONTAINS '[HSW-uWS-CT]' ) OR  [System.AreaPath] UNDER 'Task\SW\uInnovation'  OR  [System.AreaPath] UNDER 'Task\SW\MCSF1'  OR  [System.NodeName] = 'Report' ) AND  [System.WorkItemType] = 'Task'  AND  [Microsoft.VSTS.Common.Activity] = 'Improvement'  AND  [UI.Bug.ExpectFixedBranch] CONTAINS '[SWBU_69_SP4]'  AND  [System.NodeName] IN ('Heart', 'Liver Evaluation', 'Lung Nodule', 'Dual Energy Analysis', 'Dental Application', 'Colon Analysis', 'Body Perfusion', 'Bone Structure Analysis', 'Lung Density Analysis', 'Report', 'Review 4D', 'Dynamic Analysis_CT', 'Vessel Analysis_CT', 'Vessel Analysis_MR', 'Heart_Common', 'Vessel_Common', 'Vessel_Heart_Combined', 'FFR', 'Brain Perfusion_CT_3D')  AND  [System.State] IN ('Assigned', 'New')  AND  [UI.ExpectedSolvedDate] < @today  AND  [System.AssignedTo] NOT CONTAINS 'CCB'  AND  [UI.Bug.ValidatedBranch] NOT CONTAINS '[SWBU_69_SP4]' ORDER BY [System.State], [System.NodeName]";
            string sQueryTaskUnReviewed = @"SELECT [System.Id], [System.WorkItemType], [System.Title], [Microsoft.VSTS.Common.Priority], [System.AssignedTo], [System.State], [UI.bug.keyword], [UI.Reserved], [System.NodeName], [UI.ExpectedSolvedDate], [System.CreatedDate] FROM WorkItems WHERE [System.TeamProject] = 'Task'  AND (( [System.AreaPath] UNDER 'Task\SW\UIDeal_AWS'  AND  [MR.Requirement.ApplicativeProject] CONTAINS '[HSW-uWS-CT]' ) OR  [System.AreaPath] UNDER 'Task\SW\uInnovation'  OR  [System.AreaPath] UNDER 'Task\SW\MCSF1'  OR  [System.NodeName] = 'Report' ) AND  [System.WorkItemType] = 'Task'  AND  [Microsoft.VSTS.Common.Activity] = 'Improvement'  AND ( [System.State] = 'New'  OR  [System.State] = 'Assigned' ) AND  [System.NodeName] IN ('Heart', 'Liver Evaluation', 'Lung Nodule', 'Dual Energy Analysis', 'Dental Application', 'Colon Analysis', 'Body Perfusion', 'Bone Structure Analysis', 'Lung Density Analysis', 'Report', 'Review 4D', 'Dynamic Analysis_CT', 'Vessel Analysis_CT', 'Vessel Analysis_MR', 'Heart_Common', 'Vessel_Common', 'Vessel_Heart_Combined', 'FFR', 'Brain Perfusion_CT_3D')  AND  [UI.Bug.ExpectFixedBranch] NOT CONTAINS '[SWBU_69_SP4]'  AND  [UI.Bug.ExpectFixedBranch] NOT CONTAINS '[SWBU_71]'  AND  [UI.bug.keyword] NOT CONTAINS '[双能BU相关]'  AND  [System.AssignedTo] NOT CONTAINS 'CCB' ORDER BY [System.State], [System.NodeName]";
            string sQueryExpiredKey = @"SELECT [System.Id], [System.Title], [Microsoft.VSTS.Common.Priority], [System.AssignedTo], [System.State], [MR.Requirement.ApplicativeProject], [System.NodeName], [UI.bug.keyword], [UI.Reserved], [UI.Bug.ExpectFixedBranch], [UI.ExpectedSolvedDate], [Microsoft.VSTS.Common.ResolvedDate], [Microsoft.VSTS.Common.ResolvedBy], [System.CreatedDate], [System.ChangedDate] FROM WorkItems WHERE [System.TeamProject] = 'Task'  AND (( [System.AreaPath] UNDER 'Task\SW\UIDeal_AWS'  AND  [MR.Requirement.ApplicativeProject] CONTAINS '[HSW-uWS-CT]' ) OR  [System.AreaPath] UNDER 'Task\SW\uInnovation'  OR  [System.AreaPath] UNDER 'Task\SW\MCSF1'  OR  [System.NodeName] = 'Report' ) AND  [System.WorkItemType] = 'Task'  AND  [Microsoft.VSTS.Common.Activity] = 'Design'  AND  [UI.Bug.ExpectFixedBranch] CONTAINS '[SWBU_69_SP4]'  AND  [System.NodeName] IN ('Heart', 'Liver Evaluation', 'Lung Nodule', 'Dual Energy Analysis', 'Dental Application', 'Colon Analysis', 'Body Perfusion', 'Bone Structure Analysis', 'Lung Density Analysis', 'Report', 'Review 4D', 'Dynamic Analysis_CT', 'Vessel Analysis_CT', 'Vessel Analysis_MR', 'Heart_Common', 'Vessel_Common', 'Vessel_Heart_Combined', 'FFR', 'Brain Perfusion_CT_3D')  AND  [System.State] IN ('New', 'Assigned')  AND  [UI.bug.keyword] CONTAINS 'H2'  AND  [UI.ExpectedSolvedDate] < @today ORDER BY [System.State], [System.NodeName]";

            tfsTracker.InitializeTFS();

            tfsTracker.ExtractTaskList(sQueryTaskExpired);
            tfsTracker.WriteExcelItemList("Improvement Task Expired");

            tfsTracker.ExtractTaskList(sQueryTaskUnReviewed);
            tfsTracker.WriteExcelItemList("Improvement Task Unreviewed");

            tfsTracker.ExtractTaskList(sQueryExpiredKey);
            tfsTracker.WriteExcelItemList("Designed Task Expired");

            tfsTracker.WriteName2File(args[1]);
        }
    }
}
