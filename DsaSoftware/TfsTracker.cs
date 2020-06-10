using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.Common;
using Microsoft.Office.Interop.Excel;
using Microsoft.TeamFoundation.VersionControl.Client;
using System.Reflection;

namespace TestTFS
{
    class TfsTracker
    {
        public string UserName { get; set; }

        public string Password { get; set; }

        public string FileName { get; set; }

        private WorkItemStore _workItemStore;

        private IEnumerable<WorkItem> _workItemResult;

        private Dictionary<string, UserRequirementState> _UrResult;

        private Dictionary<string, TaskState> _taskResult;

        private List<ItemInfo> _ItemInfo;

        private List<ResolveInfo> _ResolveResult;

        private readonly Application _xlApp;

        private readonly Workbook _xlsWorkbook;

        public TfsTracker()
        {
            _xlApp = new Application
            {
                DisplayAlerts = false,
                Visible = false,
                ScreenUpdating = false
            };

            _xlsWorkbook = _xlApp.Workbooks.Add(true);

            _xlsWorkbook.Worksheets.Add(Missing.Value);
            _xlsWorkbook.Worksheets.Add(Missing.Value);
            _xlsWorkbook.Worksheets.Add(Missing.Value);
            _xlsWorkbook.Worksheets.Add(Missing.Value);
            _xlsWorkbook.Worksheets.Add(Missing.Value);
            _xlsWorkbook.Worksheets.Add(Missing.Value);
            _xlsWorkbook.Worksheets.Add(Missing.Value);
        }

        ~TfsTracker()
        {
            _xlsWorkbook.SaveAs(FileName);
            _xlsWorkbook.Close();
            _xlApp.Quit();
        }

        public void InitializeTFS()
        {
            try
            {
                // Method intentionally left empty.
                Microsoft.TeamFoundation.Client.TfsTeamProjectCollection server = new TfsTeamProjectCollection(
                    new Uri("http://tfs.united-imaging.com:8080/tfs/defaultcollection"),
                    new NetworkCredential(UserName, Password));

                _workItemStore = server.GetService<WorkItemStore>();
            }
            catch (Microsoft.TeamFoundation.TeamFoundationServerUnauthorizedException)
            {
                System.Console.WriteLine("Incorret User or Password.");
            }
        }

        public void SortItem()
        {
            _ItemInfo.Sort(
                (ItemInfo x, ItemInfo y) => x.NodeName.CompareTo(y.NodeName));
        }

        public void ExtractURQueryInfo(string sQuery)
        {
            _workItemResult = _workItemStore.Query(sQuery).Cast<WorkItem>();

            _UrResult = new Dictionary<string, UserRequirementState>();
            _UrResult.OrderBy(x => x.Key);

            foreach (WorkItem wi in _workItemResult)
            {
                string sNodeName = wi.NodeName;
                if (!_UrResult.Keys.Contains(sNodeName))
                {
                    _UrResult.Add(sNodeName, new UserRequirementState());
                }
                switch (wi.State)
                {
                    case "10-Requirement":
                    case "20-Solution":
                    case "30-Development":
                        {
                            _UrResult[sNodeName].ToBeDeveloped += 1;
                            break;
                        }
                    case "35-Resolved":
                        {
                            _UrResult[sNodeName].ToBeVerified += 1;
                            break;
                        }
                    case "40-SSIT Done":
                    case "50-SI":
                    case "60-SIT":
                        {
                            _UrResult[sNodeName].Verified += 1;
                            break;
                        }
                }
            }

            foreach (var item in _UrResult)
            {
                item.Value.TotalNumber = item.Value.ToBeDeveloped + item.Value.ToBeVerified + item.Value.Verified;
                double dPercentage = Convert.ToDouble(item.Value.ToBeVerified + item.Value.Verified) / Convert.ToDouble(item.Value.TotalNumber);
                item.Value.DevelopPercentage = String.Format("{0:P0}", dPercentage);
            }
        }

        public void WriteUR2Excel()
        {
            try
            {
                Worksheet sheet = _xlsWorkbook.Worksheets[1];

                int i = 1;
                foreach (var item in _UrResult)
                {
                    sheet.Cells[i, 1] = item.Key;
                    sheet.Cells[i, 2] = item.Value.ToBeDeveloped.ToString();
                    sheet.Cells[i, 3] = item.Value.ToBeVerified.ToString();
                    sheet.Cells[i, 4] = item.Value.Verified.ToString();
                    sheet.Cells[i, 5] = item.Value.TotalNumber.ToString();
                    sheet.Cells[i, 6] = item.Value.DevelopPercentage;
                    i += 1;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void ExtractTaskQueryInfo(string sQuery)
        {
            _workItemResult = _workItemStore.Query(sQuery).Cast<WorkItem>();

            _taskResult = new Dictionary<string, TaskState>();
            _taskResult.OrderBy(x => x.Key);

            foreach (WorkItem wi in _workItemResult)
            {
                string sNodeName = wi.NodeName;
                if (!_taskResult.Keys.Contains(sNodeName))
                {
                    _taskResult.Add(sNodeName, new TaskState());
                }
                switch (wi.State)
                {
                    case "New":
                    case "Assigned":
                        {
                            _taskResult[sNodeName].Assigned += 1;
                            break;
                        }
                    case "Resolved":
                        {
                            _taskResult[sNodeName].Resolved += 1;
                            break;
                        }
                    case "Verified":
                    case "Closed":
                        {
                            _taskResult[sNodeName].Verified += 1;
                            break;
                        }
                }
            }

            foreach (var item in _taskResult)
            {
                item.Value.Total = item.Value.Assigned + item.Value.Resolved + item.Value.Verified;
                double dPercentage = Convert.ToDouble(item.Value.Resolved + item.Value.Verified) / Convert.ToDouble(item.Value.Total);
                item.Value.Percentage = String.Format("{0:P0}", dPercentage);
            }
        }

        public void WriteTask2Excel()
        {
            try
            {
                Worksheet sheet = _xlsWorkbook.Worksheets[2];

                int i = 1;
                foreach (var item in _taskResult)
                {
                    sheet.Cells[i, 1] = item.Key;
                    sheet.Cells[i, 2] = item.Value.Assigned.ToString();
                    sheet.Cells[i, 3] = item.Value.Resolved.ToString();
                    sheet.Cells[i, 4] = item.Value.Verified.ToString();
                    sheet.Cells[i, 5] = item.Value.Total.ToString();
                    sheet.Cells[i, 6] = item.Value.Percentage;
                    i += 1;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void ExtractURList(string sQuery)
        {
            _workItemResult = _workItemStore.Query(sQuery).Cast<WorkItem>();

            _ItemInfo = new List<ItemInfo>();

            foreach (var wi in _workItemResult)
            {
                string sAssignedTo = wi["Assigned To"].ToString();
                _ItemInfo.Add(new ItemInfo()
                {
                    ID = wi.Id,
                    Title = wi.Title,
                    ExpectedSolvedDate = (wi["Finish Date"] == null) ? "" : wi["Finish Date"].ToString(),
                    NodeName = wi.NodeName,
                    AssignedTo = sAssignedTo.Substring(0, sAssignedTo.IndexOf('_'))
                });
            }
        }

        public void ExtractTaskList(string sQuery)
        {
            _workItemResult = _workItemStore.Query(sQuery).Cast<WorkItem>();

            _ItemInfo = new List<ItemInfo>();

            foreach (var wi in _workItemResult)
            {
                string sAssignedTo = wi["Assigned To"].ToString();
                _ItemInfo.Add(new ItemInfo()
                {
                    ID = wi.Id,
                    Title = wi.Title,
                    ExpectedSolvedDate = (wi["Expected Solved Date"] == null) ? "" : wi["Expected Solved Date"].ToString(),
                    NodeName = wi.NodeName,
                    AssignedTo = sAssignedTo.Substring(0, sAssignedTo.IndexOf('_'))
                });
            }
        }

        public void WriteExcelIndex(int n)
        {
            try
            {
                Worksheet sheet = _xlsWorkbook.Worksheets[n];

                int i = 1;
                foreach (var item in _ItemInfo)
                {
                    sheet.Cells[i, 1] = item.ID;
                    sheet.Cells[i, 2] = item.Title;
                    sheet.Cells[i, 3] = item.ExpectedSolvedDate;
                    sheet.Cells[i, 4] = item.NodeName;
                    sheet.Cells[i, 5] = item.AssignedTo;
                    i += 1;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void ExtractResolveInfo(string sQuery)
        {
            _workItemResult = _workItemStore.Query(sQuery).Cast<WorkItem>();

            _ResolveResult = new List<ResolveInfo>();

            foreach (var wi in _workItemResult)
            {
                string sResolvedBy = wi["Resolved By"].ToString();
                _ResolveResult.Add(new ResolveInfo()
                {
                    ID = wi.Id,
                    Title = wi.Title,
                    ResolvedDate = Convert.ToDateTime(wi["Resolved Date"].ToString(), new System.Globalization.DateTimeFormatInfo()),
                    NodeName = wi.NodeName,
                    ResolvedBy = sResolvedBy.Substring(0, sResolvedBy.IndexOf('_'))
                });
            }

            _ResolveResult.Sort((ResolveInfo x, ResolveInfo y) => x.ResolvedDate.CompareTo(y.ResolvedDate));
        }

        public void WriteResolveIndex(int n)
        {
            try
            {
                Worksheet sheet = _xlsWorkbook.Worksheets[n];

                int i = 1;
                foreach (var item in _ResolveResult)
                {
                    sheet.Cells[i, 1] = item.ID;
                    sheet.Cells[i, 2] = item.Title;
                    sheet.Cells[i, 3] = item.ResolvedDate;
                    sheet.Cells[i, 4] = item.NodeName;
                    sheet.Cells[i, 5] = item.ResolvedBy;
                    i += 1;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
