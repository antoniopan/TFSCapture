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
using System.Xml;
using System.Text;
using System.IO;
using Microsoft.TeamFoundation.Client.Reporting;

namespace TestTFS
{
    class TfsTracker
    {
        public string UserName { get; set; }

        public string Password { get; set; }

        public string FileName { get; set; }

        public string FileNameModule { get; set; }

        private WorkItemStore _workItemStore;

        private Dictionary<string, Dictionary<string, UserRequirementState>> _UrbyModule; // NodeName --> uModule --> UR State

        //private IEnumerable<WorkItem> _workItemResult;

        private Dictionary<string, UserRequirementState> _UrResult; // Convert UR Query into Customized data structure in table form

        private Dictionary<string, TaskState> _taskResult; // Convert Task Query into Customized data structure in table form

        private List<ItemInfo> _ItemInfo; // Convert either UR or Task Query into a list, usually to be resolved

        private List<ResolveInfo> _ResolveResult; // Convert either UR or Task Query into a list, usually resolved

        private readonly Application _xlApp;

        private readonly Workbook _xlsWorkbook;

        private readonly Workbook _xlsWorkbookModule;

        private readonly HashSet<string> _nameList;

        public TfsTracker()
        {
            _xlApp = new Application
            {
                DisplayAlerts = false,
                Visible = false,
                ScreenUpdating = false
            };

            _xlsWorkbook = _xlApp.Workbooks.Add(true);

            _xlsWorkbookModule = _xlApp.Workbooks.Add(true);

            _nameList = new HashSet<string>();
        }

        ~TfsTracker()
        {
            _xlsWorkbook.SaveAs(FileName);
            _xlsWorkbook.Close();
            _xlsWorkbookModule.SaveAs(FileNameModule);
            _xlsWorkbookModule.Close();
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
            var workItemResult = _workItemStore.Query(sQuery).Cast<WorkItem>();

            _UrResult = new Dictionary<string, UserRequirementState>();
            _UrbyModule = new Dictionary<string, Dictionary<string, UserRequirementState>>();

            foreach (WorkItem wi in workItemResult)
            {
                // Process _UrResult
                string sNodeName = wi.NodeName;
                if (!_UrResult.Keys.Contains(sNodeName))
                {
                    _UrResult.Add(sNodeName, new UserRequirementState());
                }

                // Process _UrbyModule
                if (!_UrbyModule.Keys.Contains(sNodeName))
                {
                    _UrbyModule.Add(sNodeName, new Dictionary<string, UserRequirementState>());
                }
                string sUModule = wi["UModule"].ToString();
                if (!_UrbyModule[sNodeName].Keys.Contains(sUModule))
                {
                    _UrbyModule[sNodeName].Add(sUModule, new UserRequirementState());
                }

                switch (wi.State)
                {
                    case "10-Requirement":
                    case "20-Solution":
                    case "30-Development":
                        {
                            _UrResult[sNodeName].ToBeDeveloped += 1;
                            _UrbyModule[sNodeName][sUModule].ToBeDeveloped += 1;
                            break;
                        }
                    case "35-Resolved":
                        {
                            _UrResult[sNodeName].ToBeVerified += 1;
                            _UrbyModule[sNodeName][sUModule].ToBeVerified += 1;
                            break;
                        }
                    case "40-SSIT Done":
                    case "50-SI":
                    case "60-SIT":
                        {
                            _UrResult[sNodeName].Verified += 1;
                            _UrbyModule[sNodeName][sUModule].Verified += 1;
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

        public void WriteUR2Excel(string sSheetName)
        {
            try
            {
                _xlsWorkbook.Worksheets.Add(Missing.Value);
                _xlsWorkbook.ActiveSheet.Name = sSheetName;
                Worksheet sheet = _xlsWorkbook.Worksheets[sSheetName];

                int i = 1;
                foreach (var item in _UrResult)
                {
                    sheet.Cells[i, 1] = item.Key;
                    sheet.Cells[i, 2] = item.Value.ToBeDeveloped.ToString();
                    sheet.Cells[i, 3] = item.Value.ToBeVerified.ToString();
                    sheet.Cells[i, 4] = item.Value.Verified.ToString();
                    sheet.Cells[i, 5] = item.Value.TotalNumber.ToString();
                    sheet.Cells[i, 6] = item.Value.DevelopPercentage;
                    i++;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void WriteUrByModule()
        {
            while (_UrbyModule.Count > _xlsWorkbookModule.Worksheets.Count)
            {
                _xlsWorkbookModule.Worksheets.Add(Missing.Value);
            }

            int i = 1;
            foreach (var urModule in _UrbyModule)
            {
                var sheet = _xlsWorkbookModule.Worksheets[i++];
                sheet.Name = urModule.Key;
                
                int j = 1;
                foreach (var item in urModule.Value)
                {
                    item.Value.TotalNumber = item.Value.ToBeDeveloped + item.Value.ToBeVerified + item.Value.Verified;
                    double dPercentage = Convert.ToDouble(item.Value.ToBeVerified + item.Value.Verified) / Convert.ToDouble(item.Value.TotalNumber);
                    item.Value.DevelopPercentage = String.Format("{0:P0}", dPercentage);

                    sheet.Cells[j, 1] = item.Key;
                    sheet.Cells[j, 2] = item.Value.ToBeDeveloped.ToString();
                    sheet.Cells[j, 3] = item.Value.ToBeVerified.ToString();
                    sheet.Cells[j, 4] = item.Value.Verified.ToString();
                    sheet.Cells[j, 5] = item.Value.TotalNumber.ToString();
                    sheet.Cells[j, 6] = item.Value.DevelopPercentage;

                    j++;
                }
            }
        }

        public void ExtractTaskQueryInfo(string sQuery)
        {
            var workItemResult = _workItemStore.Query(sQuery).Cast<WorkItem>();

            _taskResult = new Dictionary<string, TaskState>();

            foreach (WorkItem wi in workItemResult)
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
                    case "Observation":
                    case "Reject":
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
                        //{
                        //    _taskResult[sNodeName].Other += 1;
                        //    break;
                        //}
                }
            }

            foreach (var item in _taskResult)
            {
                item.Value.Total = item.Value.Assigned + item.Value.Resolved + item.Value.Verified + item.Value.Other;
                double dPercentage = Convert.ToDouble(item.Value.Resolved + item.Value.Verified) / Convert.ToDouble(item.Value.Total);
                item.Value.Percentage = String.Format("{0:P0}", dPercentage);
            }
        }

        public void WriteTask2Excel(string sSheetName)
        {
            try
            {
                _xlsWorkbook.Worksheets.Add(Missing.Value);
                _xlsWorkbook.ActiveSheet.Name = sSheetName;
                Worksheet sheet = _xlsWorkbook.Worksheets[sSheetName];

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
            var workItemResult = _workItemStore.Query(sQuery).Cast<WorkItem>();

            _ItemInfo = new List<ItemInfo>();

            foreach (var wi in workItemResult)
            {
                string sAssignedTo = wi["Assigned To"].ToString();
                _nameList.Add(sAssignedTo);
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
            var workItemResult = _workItemStore.Query(sQuery).Cast<WorkItem>();

            _ItemInfo = new List<ItemInfo>();

            foreach (var wi in workItemResult)
            {
                string sAssignedTo = wi["Assigned To"].ToString();
                _nameList.Add(sAssignedTo);
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

        public void WriteExcelItemList(string sSheetName)
        {
            try
            {
                _xlsWorkbook.Worksheets.Add(Missing.Value);
                _xlsWorkbook.ActiveSheet.Name = sSheetName;
                Worksheet sheet = _xlsWorkbook.Worksheets[sSheetName];

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
            var workItemResult = _workItemStore.Query(sQuery).Cast<WorkItem>();

            _ResolveResult = new List<ResolveInfo>();

            foreach (var wi in workItemResult)
            {
                string sResolvedBy = wi["Resolved By"].ToString();
                _nameList.Add(sResolvedBy);
                _ResolveResult.Add(new ResolveInfo()
                {
                    ID = wi.Id,
                    Title = wi.Title,
                    ResolvedDate = Convert.ToDateTime(wi["Resolved Date"].ToString(), new System.Globalization.DateTimeFormatInfo()),
                    NodeName = wi.NodeName,
                    ResolvedBy = sResolvedBy.Substring(0, sResolvedBy.IndexOf('_')),
                    Reserved = (wi.Type.Name == "Task") ? wi["Reserved"].ToString() : wi["uAttribute"].ToString()
                });
            }

            _ResolveResult.Sort((ResolveInfo x, ResolveInfo y) => x.ResolvedDate.CompareTo(y.ResolvedDate));
        }

        public void WriteResolveItemList(string sSheetName)
        {
            try
            {
                _xlsWorkbook.Worksheets.Add(Missing.Value);
                _xlsWorkbook.ActiveSheet.Name = sSheetName;
                Worksheet sheet = _xlsWorkbook.Worksheets[sSheetName];

                int i = 1;
                foreach (var item in _ResolveResult)
                {
                    sheet.Cells[i, 1] = item.ID;
                    sheet.Cells[i, 2] = item.Title;
                    sheet.Cells[i, 3] = item.ResolvedDate;
                    sheet.Cells[i, 4] = item.NodeName;
                    sheet.Cells[i, 5] = item.ResolvedBy;
                    sheet.Cells[i, 6] = item.Reserved;
                    i += 1;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void WriteName2File(string sFilename)
        {
            var nameArray = _nameList.ToArray();

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < nameArray.Length-1; i++)
            {
                sb.Append(nameArray[i]);
                sb.Append(';');
            }
            if (nameArray.Length != 0)
            {
                sb.Append(nameArray.Last());
            }

            var sTotal = sb.ToString();

            StreamWriter sw = new StreamWriter(sFilename);
            sw.WriteLine(sTotal);
            sw.Flush();
            sw.Close();
        }
    }
}
