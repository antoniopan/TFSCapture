using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestTFS
{
    class UserRequirementState
    {
        public int ToBeDeveloped { get; set; }

        public int ToBeVerified { get; set; }

        public int Verified { get; set; }

        public int TotalNumber { get; set; }

        public string DevelopPercentage { get; set; }

        public UserRequirementState()
        {
            ToBeDeveloped = 0;
            ToBeVerified = 0;
            Verified = 0;
        }
    }

    class TaskState
    {
        public int Assigned { get; set; }

        public int Resolved { get; set; }

        public int Verified { get; set; }

        public int Total { get; set; }

        public string Percentage { get; set; }

        public TaskState()
        {
            Assigned = 0;
            Resolved = 0;
            Verified = 0;
        }
    }

    class ItemInfo
    {
        public int ID { get; set; }

        public string Title { get; set; }

        public string ExpectedSolvedDate { get; set; }

        public string NodeName { get; set; }

        public string AssignedTo { get; set; }
    }

    class ResolveInfo
    {
        public int ID { get; set; }

        public string Title { get; set; }

        public string NodeName { get; set; }

        public string ResolvedBy { get; set; }

        public DateTime ResolvedDate { get; set; }

        public string Reserved { get; set; }
    }
}
