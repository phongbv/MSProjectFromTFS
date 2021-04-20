using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static System.Console;
using System.Threading.Tasks;
using System.Collections.Concurrent;
#pragma warning disable CS0618 // Type or member is obsolete
namespace MSProjectFromTFS
{

    public class TFSUtil
    {
        static WorkItemStore workItemStore;

        static string teamProjectName = "iLendingPro";
        static Uri tfsUri = new Uri("http://sptserver.ists.corp:8080/tfs/" + teamProjectName);
        static TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(tfsUri);
        static VersionControlServer vcs;
        static TFSUtil()
        {
            tpc.Authenticate();
            vcs = tpc.GetService<VersionControlServer>();
            workItemStore = new WorkItemStore(tpc);
        }
        public List<WorkItemInfo> GetNewWit()
        {
            string query = @"SELECT
        [System.Id],
        [System.WorkItemType],
        [System.Title],
        [System.State],
        [System.AreaPath],
        [System.IterationPath]
FROM workitems
WHERE
        [System.TeamProject] = @project
        and [System.WorkItemType] in(""Product Backlog Item"", ""Customer Backlog Item"")
       AND [SYStem.State] in (""New"",""InProgress"", ""In Progress"", ""Transfer Requirement"",""Response"")
       AND [System.IterationPath] NOT IN(""iLendingPro\LVPB - Performance Test"", ""iLendingPro\NCB - Credit Rating"", ""iLendingPro\Next Release"", ""iLendingPro\VIB Demo"", ""iLendingPro\LOS - Version 2.0"")
ORDER BY[System.IterationPath], [SYStem.State], [System.ChangedDate] DESC";
            Dictionary<string, string> variables = new Dictionary<string, string> { { "project", teamProjectName } };
            var workItemColl = workItemStore.Query(query, variables).OfType<WorkItem>().ToList();

            var lstWorkItem = workItemColl.Select(e => new WorkItemInfo(workItemStore) { WorkItem = e }).ToList();
            return lstWorkItem;
        }

        public WorkItemInfo GetWorkItem(int id)
        {
            return new WorkItemInfo(workItemStore)
            {
                WorkItem = workItemStore.GetWorkItem(id)
            };
        }
    }
    public class WorkItemInfo
    {
        public WorkItemInfo(WorkItemStore workItemStore)
        {
            this.workItemStore = workItemStore;
        }
        Task calculateRelatedItem;
        private WorkItem _workItem;

        WorkItemStore workItemStore;
        WorkItem featureWIT = null;
        public WorkItem WorkItem
        {
            get => _workItem;
            set
            {
                _workItem = value;
                calculateRelatedItem = Task.Run(() =>
                {
                    //WriteLine($"Reading {WorkItem.Id}");
                    _dependItem = WorkItem.Links.OfType<RelatedLink>().Where(e => e.LinkTypeEnd.Name == "Predecessor").ToList();
                    //WriteLine($"Reading {WorkItem.Id}");
                    _childItem = WorkItem.Links.OfType<RelatedLink>().Where(e => e.LinkTypeEnd.Name == "Child").ToList();
                    foreach (var parentLink in WorkItem.Links.OfType<RelatedLink>().Where(e => e.LinkTypeEnd.Name == "Parent"))
                        if (parentLink != null)
                        {
                            var wit = workItemStore.GetWorkItem(parentLink.RelatedWorkItemId);
                            if (wit.Type.Name == "Feature")
                            {
                                featureWIT = wit;
                            }
                        }
                });
            }
        }

        private List<RelatedLink> _dependItem = new List<RelatedLink>();
        public List<int> DependItemId
        {
            get
            {
                if (calculateRelatedItem != null)
                {
                    Task.WaitAll(calculateRelatedItem);
                    calculateRelatedItem = null;
                }
                return _dependItem.Select(e => e.RelatedWorkItemId).ToList();
            }
        }
        private List<RelatedLink> _childItem = new List<RelatedLink>();
        public List<int> ChildItemId
        {
            get
            {
                if (calculateRelatedItem != null)
                {
                    Task.WaitAll(calculateRelatedItem);
                    calculateRelatedItem = null;
                }
                return _childItem.Select(e => e.RelatedWorkItemId).ToList();
            }
        }
        public string IterationPath => _workItem?.IterationPath;
        public string HandleBy => _workItem?.Fields["Handle By"]?.Value + "";
        public string AreaPath
        {
            get
            {
                if (calculateRelatedItem != null)
                {
                    Task.WaitAll(calculateRelatedItem);
                    calculateRelatedItem = null;
                }
                return featureWIT?.Title ?? "Chưa gắn feature";
            }
        }
        private static List<string> StateComplete = new List<string>()
        {
            "Committed",
            "Completed"
        };
        public bool IsComplete => StateComplete.Contains(_workItem.State);
    }

}
#pragma warning restore CS0618 // Type or member is obsolete