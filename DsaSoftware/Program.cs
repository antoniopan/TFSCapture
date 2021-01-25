using System;
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
            var tfsTracker = new TFS_TRACKER.TfsTracker()
            {
                UserName = "liangliang.pan",
                Password = "4Antonio",
                FileName = args[0],
                FileNameModule = "haha"
            };

            var queries = TFS_TRACKER.TfsTracker.Deserialize(args[2]);
            TFS_TRACKER.TfsTracker.ProcessQueryXml(queries);

            tfsTracker.InitializeTFS();

            foreach (var query in queries.query)
            {
                switch (query.parsemethod)
                {
                    case queryTypeParsemethod.resolvelist:
                        tfsTracker.ExtractResolveInfo(query.queryinfo);
                        tfsTracker.WriteResolveItemList(query.queryname);
                        break;
                    case queryTypeParsemethod.tasklist:
                        tfsTracker.ExtractTaskList(query.queryinfo);
                        tfsTracker.WriteExcelItemList(query.queryname);
                        break;
                    case queryTypeParsemethod.urlist:
                        tfsTracker.ExtractURList(query.queryinfo);
                        tfsTracker.WriteExcelItemList(query.queryname);
                        break;
                    case queryTypeParsemethod.tasktable:
                        tfsTracker.ExtractTaskQueryInfo(query.queryinfo);
                        tfsTracker.WriteTask2Excel(query.queryname);
                        break;
                    case queryTypeParsemethod.urtable:
                        tfsTracker.ExtractURQueryInfo(query.queryinfo);
                        tfsTracker.WriteUR2Excel(query.queryname);
                        break;
                    case queryTypeParsemethod.createresolve:
                        tfsTracker.ExtractCreateResolve(query.queryinfo, query.additionalqueryinfo);
                        tfsTracker.WriteCreateResolveInfo(query.queryname);
                        break;
                }
            }

            tfsTracker.WriteName2File(args[1]);
        }
    }
}
