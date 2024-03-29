﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QueryCapture
{
    static class Program
    {
        static void Main(string[] args)
        {
            var queries = TFS_TRACKER.TfsTracker.Deserialize(args[2]);
            TFS_TRACKER.TfsTracker.ProcessQueryXml(queries);

            var tfsTracker = new TFS_TRACKER.TfsTracker()
            {
                UserName = queries.user.username,
                Password = queries.user.password,
                FileName = args[0],
                FileNameModule = "haha"
            };

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
