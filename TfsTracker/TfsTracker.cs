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

namespace TFS_TRACKER
{
    public class TfsTracker
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

        private Dictionary<string, int[]> _CreateResolveInfo;

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
                        {
                            _UrResult[sNodeName].ToBeReviewed += 1;
                            _UrbyModule[sNodeName][sUModule].ToBeReviewed += 1;
                            break;
                        }
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
                item.Value.TotalNumber = item.Value.ToBeReviewed + item.Value.ToBeDeveloped + item.Value.ToBeVerified + item.Value.Verified;
                double dPercentage = Convert.ToDouble(item.Value.ToBeVerified + item.Value.Verified) / Convert.ToDouble(item.Value.TotalNumber);
                item.Value.DevelopPercentage = String.Format("{0:P0}", dPercentage);
            }

            _UrResult = _UrResult.OrderBy(p => p.Key).ToDictionary(p => p.Key, o => o.Value);
            _UrbyModule = _UrbyModule.OrderBy(p => p.Key).ToDictionary(p => p.Key, o => o.Value);
        }

        public void WriteUR2Excel(string sSheetName)
        {
            try
            {
                _xlsWorkbook.Worksheets.Add(Missing.Value);
                _xlsWorkbook.ActiveSheet.Name = sSheetName;
                Worksheet sheet = _xlsWorkbook.Worksheets[sSheetName];

                int i = 1;
                var summaryUR = new UserRequirementState();
                foreach (var item in _UrResult)
                {
                    sheet.Cells[i, 1] = item.Key;
                    sheet.Cells[i, 2] = item.Value.ToBeReviewed.ToString();
                    sheet.Cells[i, 3] = item.Value.ToBeDeveloped.ToString();
                    sheet.Cells[i, 4] = item.Value.ToBeVerified.ToString();
                    sheet.Cells[i, 5] = item.Value.Verified.ToString();
                    sheet.Cells[i, 6] = item.Value.TotalNumber.ToString();
                    sheet.Cells[i, 7] = item.Value.DevelopPercentage;
                    summaryUR.ToBeReviewed += item.Value.ToBeReviewed;
                    summaryUR.ToBeDeveloped += item.Value.ToBeDeveloped;
                    summaryUR.ToBeVerified += item.Value.ToBeVerified;
                    summaryUR.Verified += item.Value.Verified;
                    summaryUR.TotalNumber += item.Value.TotalNumber;
                    i++;
                }

                sheet.Cells[i, 1] = @"总计";
                sheet.Cells[i, 2] = summaryUR.ToBeReviewed.ToString();
                sheet.Cells[i, 3] = summaryUR.ToBeDeveloped.ToString();
                sheet.Cells[i, 4] = summaryUR.ToBeVerified.ToString();
                sheet.Cells[i, 5] = summaryUR.Verified.ToString();
                sheet.Cells[i, 6] = summaryUR.TotalNumber.ToString();
                double dPercentage = Convert.ToDouble(summaryUR.ToBeVerified + summaryUR.Verified) / Convert.ToDouble(summaryUR.TotalNumber);
                sheet.Cells[i, 7] = String.Format("{0:P0}", dPercentage);
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
                var summaryUR = new UserRequirementState();
                foreach (var item in urModule.Value)
                {
                    item.Value.TotalNumber = item.Value.ToBeReviewed + item.Value.ToBeDeveloped + item.Value.ToBeVerified + item.Value.Verified;
                    double dPercentage = Convert.ToDouble(item.Value.ToBeVerified + item.Value.Verified) / Convert.ToDouble(item.Value.TotalNumber);
                    item.Value.DevelopPercentage = String.Format("{0:P0}", dPercentage);

                    sheet.Cells[j, 1] = item.Key;
                    sheet.Cells[j, 2] = item.Value.ToBeReviewed.ToString();
                    sheet.Cells[j, 3] = item.Value.ToBeDeveloped.ToString();
                    sheet.Cells[j, 4] = item.Value.ToBeVerified.ToString();
                    sheet.Cells[j, 5] = item.Value.Verified.ToString();
                    sheet.Cells[j, 6] = item.Value.TotalNumber.ToString();
                    sheet.Cells[j, 7] = item.Value.DevelopPercentage;

                    summaryUR.ToBeReviewed += item.Value.ToBeReviewed;
                    summaryUR.ToBeDeveloped += item.Value.ToBeDeveloped;
                    summaryUR.ToBeVerified += item.Value.ToBeVerified;
                    summaryUR.Verified += item.Value.Verified;
                    summaryUR.TotalNumber += item.Value.TotalNumber;
                    j++;
                }

                sheet.Cells[j, 1] = @"总计";
                sheet.Cells[j, 2] = summaryUR.ToBeReviewed.ToString();
                sheet.Cells[j, 3] = summaryUR.ToBeDeveloped.ToString();
                sheet.Cells[j, 4] = summaryUR.ToBeVerified.ToString();
                sheet.Cells[j, 5] = summaryUR.Verified.ToString();
                sheet.Cells[j, 6] = summaryUR.TotalNumber.ToString();
                sheet.Cells[j, 7] = String.Format("{0:P0}", Convert.ToDouble(summaryUR.ToBeVerified + summaryUR.Verified) / Convert.ToDouble(summaryUR.TotalNumber));
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
                    case "Terminated":
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

            _taskResult = _taskResult.OrderBy(p => p.Key).ToDictionary(p => p.Key, o => o.Value);
        }

        public void WriteTask2Excel(string sSheetName)
        {
            try
            {
                _xlsWorkbook.Worksheets.Add(Missing.Value);
                _xlsWorkbook.ActiveSheet.Name = sSheetName;
                Worksheet sheet = _xlsWorkbook.Worksheets[sSheetName];

                int i = 1;
                var _summaryTask = new TaskState();
                foreach (var item in _taskResult)
                {
                    sheet.Cells[i, 1] = item.Key;
                    sheet.Cells[i, 2] = item.Value.Assigned.ToString();
                    sheet.Cells[i, 3] = item.Value.Resolved.ToString();
                    sheet.Cells[i, 4] = item.Value.Verified.ToString();
                    sheet.Cells[i, 5] = item.Value.Total.ToString();
                    sheet.Cells[i, 6] = item.Value.Percentage;
                    sheet.Cells[i, 7] = String.Format("{0:P0}", Convert.ToDouble(item.Value.Verified) / Convert.ToDouble(item.Value.Total));
                    _summaryTask.Assigned += item.Value.Assigned;
                    _summaryTask.Resolved += item.Value.Resolved;
                    _summaryTask.Verified += item.Value.Verified;
                    _summaryTask.Total += item.Value.Total;
                    i += 1;
                }
                sheet.Cells[i, 1] = @"总计";
                sheet.Cells[i, 2] = _summaryTask.Assigned.ToString();
                sheet.Cells[i, 3] = _summaryTask.Resolved.ToString();
                sheet.Cells[i, 4] = _summaryTask.Verified.ToString();
                sheet.Cells[i, 5] = _summaryTask.Total.ToString();
                double dPercentage = Convert.ToDouble(_summaryTask.Resolved + _summaryTask.Verified) / Convert.ToDouble(_summaryTask.Total);
                sheet.Cells[i, 6] = String.Format("{0:P0}", dPercentage);
                dPercentage = Convert.ToDouble(_summaryTask.Verified) / Convert.ToDouble(_summaryTask.Total);
                sheet.Cells[i, 7] = String.Format("{0:P0}", dPercentage);
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
                int iIndex = sAssignedTo.IndexOf('-');
                _nameList.Add(sAssignedTo);
                _ItemInfo.Add(new ItemInfo()
                {
                    ID = wi.Id,
                    Title = wi.Title,
                    ExpectedSolvedDate = (wi["Expected Solved Date"] == null) ? "" : wi["Expected Solved Date"].ToString(),
                    NodeName = wi.NodeName,
                    AssignedTo = iIndex > 0 ? sAssignedTo.Substring(0, sAssignedTo.IndexOf('_')) : sAssignedTo,
                    Priority = wi["Priority"].ToString()
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
                    sheet.Cells[i, 3] = item.Priority;
                    sheet.Cells[i, 4] = item.NodeName;
                    sheet.Cells[i, 5] = item.AssignedTo;
                    sheet.Cells[i, 6] = item.ExpectedSolvedDate;
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
                int iIndex = sResolvedBy.IndexOf('_');
                _nameList.Add(sResolvedBy);
                _ResolveResult.Add(new ResolveInfo()
                {
                    ID = wi.Id,
                    Title = wi.Title,
                    ResolvedDate = Convert.ToDateTime(wi["Resolved Date"].ToString(), new System.Globalization.DateTimeFormatInfo()),
                    NodeName = wi.NodeName,
                    ResolvedBy = iIndex > 0 ? sResolvedBy.Substring(0, sResolvedBy.IndexOf('_')) : "",
                    Reserved = (wi.Type.Name == "Task") ? String.Format("P{0}", wi["Priority"].ToString()) : ""
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
                    sheet.Cells[i, 3] = item.Reserved;
                    sheet.Cells[i, 4] = item.NodeName;
                    sheet.Cells[i, 5] = item.ResolvedBy;
                    sheet.Cells[i, 6] = item.ResolvedDate;
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

        public void ExtractCreateResolve(string sQueryCreate, string sQueryResolve)
        {
            _CreateResolveInfo = new Dictionary<string, int[]>();

            var workItemResult = _workItemStore.Query(sQueryCreate).Cast<WorkItem>();
            foreach (var wi in workItemResult)
            {
                if (!_CreateResolveInfo.Keys.Contains(wi.NodeName))
                {
                    _CreateResolveInfo.Add(wi.NodeName, new int[] { 0, 0 });
                }
                _CreateResolveInfo[wi.NodeName][0] += 1;
            }

            workItemResult = _workItemStore.Query(sQueryResolve).Cast<WorkItem>();
            foreach (var wi in workItemResult)
            {
                if (!_CreateResolveInfo.Keys.Contains(wi.NodeName))
                {
                    _CreateResolveInfo.Add(wi.NodeName, new int[] { 0, 0 });
                }
                _CreateResolveInfo[wi.NodeName][1] += 1;
            }

            _CreateResolveInfo = _CreateResolveInfo.OrderBy(p => p.Key).ToDictionary(p => p.Key, o => o.Value);
        }

        public void WriteCreateResolveInfo(string sSheetName)
        {
            try
            {
                _xlsWorkbook.Worksheets.Add(Missing.Value);
                _xlsWorkbook.ActiveSheet.Name = sSheetName;
                Worksheet sheet = _xlsWorkbook.Worksheets[sSheetName];

                int i = 1;
                int iCreate = 0;
                int iResolved = 0;
                foreach (var item in _CreateResolveInfo)
                {
                    sheet.Cells[i, 1] = item.Key;
                    sheet.Cells[i, 2] = item.Value[0];
                    sheet.Cells[i, 3] = item.Value[1];
                    i += 1;
                    iCreate += item.Value[0];
                    iResolved += item.Value[1];
                }
                sheet.Cells[i, 1] = @"总和";
                sheet.Cells[i, 2] = iCreate;
                sheet.Cells[i, 3] = iResolved;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public static int FindLastMonday()
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

        public static void ProcessQueryXml(projectqueryType projQuery)
        {
            foreach(var query in projQuery.query)
            {
                if (query.replacetoday)
                {
                    query.queryinfo = query.queryinfo.Replace("@today", String.Format("@today-{0:D}", FindLastMonday()));
                    if (query.additionalqueryinfo != null)
                    {
                        query.additionalqueryinfo = query.additionalqueryinfo.Replace("@today", String.Format("@today-{0:D}", FindLastMonday()));
                    }
                }
            }
        }

        public static projectqueryType Deserialize(string path)
        {
            object obj = null;

            if (File.Exists(path))
            {
                using (StreamReader sr = new StreamReader(path))
                {
                    System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(projectqueryType));
                    obj = serializer.Deserialize(sr);
                }
            }

            return obj as projectqueryType;
        }
    }
}
