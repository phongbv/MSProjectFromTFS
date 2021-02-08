using Microsoft.Office.Interop.MSProject;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace MSProjectFromTFS
{
    class Program
    {
        MSProject.Application projApp = null;
        MSProject.Projects projects = null;
        MSProject.Project project = null;
        MSProject.Tasks tasks = null;
        private static string filePath = ConfigurationManager.AppSettings["filePath"];
        TFSUtil tfsUtil = new TFSUtil();
        Dictionary<string, Task> allTasks = new Dictionary<string, Task>();
        static void Main(string[] args)
        {
            new Program().DoFillData();

        }
        List<WorkItemInfo> lstWit;
        private void DoFillData()
        {
            try
            {
                lstWit = tfsUtil.GetNewWit();
                projApp = new MSProject.Application();
                bool fileExist = File.Exists(filePath);
                if (fileExist == false)
                {
                    projects = projApp.Projects;
                    project = projects.Add(false, null, false);
                    tasks = project.Tasks;
                    projApp.Application.TableEditEx(Name: "&Entry", TaskTable: true, FieldName: "Text1", Width: 15, ColumnPosition: 0, Title: "Wit Id");
                    projApp.Application.TableEditEx(Name: "&Entry", TaskTable: true, FieldName: "Text2", Width: 20, ColumnPosition: 1, Title: "Wit Type", WrapText: true);
                    projApp.Application.TableEditEx(Name: "&Entry", TaskTable: true, FieldName: "Text3", Width: 23, ColumnPosition: 2, Title: "State");
                    projApp.Application.TableEditEx(Name: "&Entry", TaskTable: true, FieldName: "Name", ColumnPosition: 3, Width: 120);
                    projApp.Application.TableEditEx(Name: "&Entry", TaskTable: true, FieldName: "Duration", ColumnPosition: 4);
                    projApp.Application.TableEditEx(Name: "&Entry", TaskTable: true, FieldName: "Start", ColumnPosition: 5);
                    projApp.Application.TableEditEx(Name: "&Entry", TaskTable: true, FieldName: "Finish", ColumnPosition: 6);
                    projApp.Application.TableEditEx(Name: "&Entry", TaskTable: true, FieldName: "Predecessors", ColumnPosition: 7);
                    projApp.Application.TableEditEx(Name: "&Entry", TaskTable: true, NewFieldName: "Resource Names", ColumnPosition: 8, Width: 30);
                    projApp.Application.TableEditEx(Name: "&Entry", TaskTable: true, NewFieldName: "% Complete", ColumnPosition: 9, Width: 30);
                    //projApp.Application.TableEditEx(Name: "&Entry", TaskTable: true,,  FieldName: "Resource Names", ColumnPosition: 8);
                    projApp.Application.TableApply(Name: "&Entry");
                    // lstWit = lstWit.Where(e => e.WorkItem.Id != 14740).ToList();
                }
                else
                {
                    object missingValue = System.Reflection.Missing.Value;
                    projApp.FileOpenEx(filePath,
                        missingValue, missingValue, missingValue, missingValue,
                        missingValue, missingValue, missingValue, missingValue,
                        missingValue, missingValue, PjPoolOpen.pjPoolReadWrite,
                        missingValue, missingValue, missingValue, missingValue,
                        missingValue);

                    //Create a Project object by assigning active project
                    project = projApp.ActiveProject;
                    tasks = project.Tasks;
                    // Union list item mới và list cũ để update state, insert record
                    foreach (var task in tasks.OfType<Task>().Where(e => string.IsNullOrEmpty(e.Text1) == false && lstWit.Select(x => x.WorkItem.Id + "").Contains(e.Text1) == false))
                    {
                        lstWit.Add(tfsUtil.GetWorkItem(int.Parse(task.Text1)));
                    }

                }
                var tmpList = tasks.OfType<Task>().Where(e => string.IsNullOrEmpty(e.Name) == false);
                var a = tmpList.Select(e => new
                {
                    e.Text1,
                    key = e.Name + "_" + e.Text1,
                    parent = e.Text30,
                    task = e
                }).ToList();
                allTasks = a.ToDictionary(e => string.IsNullOrEmpty(e.Text1) ? e.key + "_" + e.parent : e.Text1, x => x.task);

                lstWit = lstWit.OrderBy(e => e.WorkItem.Id).ToList();

                foreach (var customTaskGroup in lstWit.GroupBy(e => e.IterationPath))
                {
                    var taskLevel1 = AddTask(customTaskGroup.Key, null, 1, null, true);

                    foreach (var witModule in customTaskGroup.ToList().GroupBy(e => e.AreaPath))
                    {
                        var taskLevel2 = AddTask(witModule.Key, null, 2, taskLevel1, true);
                        foreach (var wit in witModule)
                        {
                            AddTask(null, wit, 3, taskLevel2);
                        }
                    }
                }
                UpdatePredecessorTasks();
                RemoveEmptyTask();
                if (fileExist)
                {
                    projApp.FileCloseEx(PjSaveType.pjSave);
                }
                else
                {
                    project.SaveAs(filePath);
                }

                projApp.Quit();
            }
            finally
            {
                if (project != null) Marshal.ReleaseComObject(project);
                if (projects != null) Marshal.ReleaseComObject(projects);
                if (projApp != null) Marshal.ReleaseComObject(projApp);
            }
        }

        private void RemoveEmptyTask()
        {
            foreach (var item in tasks.OfType<Task>().Where(e => string.IsNullOrEmpty(e.Name)).ToList())
            {
                item.Delete();
            }
        }

        private void UpdatePredecessorTasks()
        {
            Dictionary<string, Task> lstTasks = tasks.OfType<Task>().Where(e => string.IsNullOrEmpty(e.Text1) == false).ToDictionary(e => e.Text1);
            foreach (var workItem in lstWit)
            {
                if (workItem.DependItemId.Any())
                {
                    var currentTask = lstTasks[workItem.WorkItem.Id + ""];
                    //currentTask.Predecessors = "";
                    foreach (var item in workItem.DependItemId)
                    {
                        //currentTask.PredecessorTasks.Add(lstTasks[item]);
                        if (lstTasks.ContainsKey(item + ""))
                            currentTask.Predecessors += lstTasks[item + ""].ID;
                    }
                }
            }
        }

        private
        Task AddTask(string taskTitle, WorkItemInfo wit, int level, Task parentTask = null, bool forceUpdateLevel = false)
        {

            if (wit != null)
            {
                taskTitle = wit.WorkItem.Title;
            }
            else if (allTasks.ContainsKey(taskTitle + "__" + parentTask?.Name)) return allTasks[taskTitle + "__" + parentTask?.Name];
            Console.WriteLine("Update task " + taskTitle + ", Level " + level);

            double? witId = wit == null ? default(double?) : wit.WorkItem.Id;
            string key = witId.HasValue ? witId.ToString() : taskTitle + "_" + witId + "_" + parentTask?.Name;
            bool isExists = allTasks.ContainsKey(key);
            var witTask = isExists ? allTasks[key] :/*tasks.Add(taskTitle)*/ InsertChildTask(parentTask, taskTitle, tasks);
            if (wit != null)
                allTasks[key] = witTask;
            if (forceUpdateLevel)
            {
                witTask.OutlineLevel = (short)level;
            }
            //if (isExists == false)
            //{
            //    witTask.OutlineLevel = (short)level;
            //}
            if (wit != null)
            {
                witTask.Name = wit.WorkItem.Title;
                witTask.Text1 = wit.WorkItem.Id + "";
                if (witTask.Start == null)
                    witTask.Start = wit.WorkItem.Fields["Dev Start Date"].Value;
                if (witTask.Finish == null)
                    witTask.Finish = wit.WorkItem.Fields["Dev Due Date"].Value;
                if (string.IsNullOrEmpty(witTask.ResourceNames))
                {
                    string assignTo = wit.WorkItem.Fields[CoreField.AssignedTo]?.Value + "";
                    assignTo = assignTo.Replace(", ", "").Replace(" (iSTS)", "");
                    var lstValues = assignTo.Split(',');
                    witTask.ResourceNames = lstValues[0] + (lstValues.Length == 2 ? String.Join("", lstValues[1].Split(' ').Select(e => e.FirstOrDefault())) : "");
                }
                else
                {
                    string assignTo = witTask.ResourceNames;
                    assignTo = assignTo.Replace(", ", "").Replace(" (iSTS)", "");
                    var lstValues = assignTo.Split(',');
                    var tmp = lstValues[0] + (lstValues.Length == 2 ? String.Join("", lstValues[1].Split(' ').Select(e => e.FirstOrDefault())) : "");
                    witTask.ResourceNames = lstValues[0] + (lstValues.Length == 2 ? String.Join("", lstValues[1].Split(' ').Select(e => e.FirstOrDefault())) : "");
                }

                witTask.Text2 = wit.WorkItem.Type.Name;
                witTask.Text3 = wit.WorkItem.Fields[CoreField.State]?.Value + "";
                if (wit.IsComplete)
                {
                    witTask.PercentComplete = 100;
                }
            }
            witTask.Number1 = witTask.OutlineLevel;
            witTask.Text30 = parentTask?.Name;
            projApp.FileSave();
            return witTask;
        }
        Task InsertChildTask(Task parentTask, string taskTitle, Tasks globalTaskList)
        {
            if (parentTask == null)
            {
                return globalTaskList.Add(taskTitle);
            }

            Task currentChildTask;
            // Lấy ra thằng con cuối cùng
            //var lastChildTask = parentTask.OutlineChildren.OfType<Task>().LastOrDefault();
            // Nếu thằng con cuối cùng ko có thì tạo mới thằng con vào outline outdent nó vào
            //if (lastChildTask == null)
            //{
            //    currentChildTask = globalTaskList.Add(taskTitle, parentTask.ID + 1);

            //    currentChildTask.OutlineIndent();
            //}
            //else
            //{
            //    currentChildTask = globalTaskList.Add(taskTitle, lastChildTask.ID + 1);
            //}
            var lastId = GetLastId(parentTask);
            currentChildTask = globalTaskList.Add(taskTitle, lastId + 1);
            switch (currentChildTask.OutlineLevel)
            {
                case 1:
                case 2:
                    if (currentChildTask.ID != 1)
                    {
                        if (parentTask.OutlineLevel == currentChildTask.OutlineLevel)
                            currentChildTask.OutlineIndent();
                    }
                    break;
                default:
                    //if (currentChildTask.ID != 1) currentChildTask.OutlineIndent();
                    break;
            }
            //currentChildTask.Parent = parentTask;
            return currentChildTask;
        }

        private int GetLastId(Task t)
        {
            if (t.OutlineChildren.Count == 0) return t.ID;
            var lastChildTask = t.OutlineChildren.OfType<Task>().LastOrDefault();
            return GetLastId(lastChildTask);
        }

    }
}
public static class EnumeratorExtensions
{
    public static IEnumerable<T> ToEnumerable<T>(this IEnumerator<T> enumerator)
    {
        while (enumerator.MoveNext())
            yield return enumerator.Current;
    }
}