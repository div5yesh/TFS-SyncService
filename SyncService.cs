using System;
using System.Data;
using System.ServiceProcess;
using System.Timers;
using System.IO;
using System.Data.SqlClient;
using System.Collections;
using System.Configuration;
using System.Net;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace TfsSyncService
{
    public partial class SyncService : ServiceBase
    {
        System.Timers.Timer timer;
        //ArrayList TeamProjectArrayList;
        public SyncService()
        {
            try
            {
                InitializeComponent();
                timer = new System.Timers.Timer();
                timer.Elapsed += new ElapsedEventHandler(TfsSync);
                int time = Convert.ToInt32(ConfigurationManager.AppSettings["Time"]);
                if (time < 1) time = 1;
                timer.Interval = time * 60 * 1000;
                timer.Enabled = true;
            }
            catch (ObjectDisposedException) { ServiceEventLog("Sync cannot be started - N1"); }
            catch (FormatException)
            {
                ServiceEventLog("Invalid time format - N1");
                timer.Stop();
                timer.Dispose();
            }
            catch (OverflowException)
            {
                ServiceEventLog("Time exceeded the limit - N1");
                timer.Stop();
                timer.Dispose();
            }
            catch (ConfigurationErrorsException) { ServiceEventLog("No valid configuration file found - N1"); }
            catch (Exception e){ ServiceEventLog("Something went wrong - N1"+e); }
        }
        protected override void OnStart(string[] args)
        {
            ServiceEventLog("Starting service");
            timer.Start();
        }
        private void TfsSync(object sender, ElapsedEventArgs e)
        {
            //GetCollections();
            //ServiceEventLog("###Synced collections###");
            //GetProjects();
            GetProjectsAndCollections();
            ServiceEventLog("###Synced projects and collections###");
            GetWorkitems();
            ServiceEventLog("###Synced workitems###");
            SyncWorkitems();
            ServiceEventLog("###Synced user workitems###");
           // SyncHistory();
           // ServiceEventLog("###Synced history###");
            ServiceEventLog("@@@Sync step completed@@@");
            ServiceEventLog("------------------------------------------------------------------");
        }
        //private void GetCollections()
        //{
        //    try
        //    {
        //        String CollectionQuery = "select ProjectNodeName from Tfs_Warehouse.dbo.DimTeamProject where Depth=0";
        //        SqlConnection cn = new SqlConnection(ConfigurationManager.AppSettings["TfsDataSrc"]);
        //        cn.Open();
        //        SqlCommand cmd = new SqlCommand(CollectionQuery, cn);
        //        SqlDataAdapter da = new SqlDataAdapter(cmd);
        //        DataSet ds = new DataSet();
        //        da.Fill(ds, "Collections");
        //        CollectionAL = new ArrayList();
        //        foreach (DataRow row in ds.Tables["Collections"].Rows)
        //        {
        //            String dbname = row["ProjectNodeName"].ToString();
        //            CollectionAL.Add(dbname);
        //            ServiceEventLog(dbname);
        //        }
        //    }
        //    catch (ConfigurationErrorsException) { ServiceEventLog("No valid configuration file found - N2"); }
        //    catch (SqlException e) { ServiceEventLog("Database access error - N2"+e); }
        //    catch (Exception e){ ServiceEventLog("Something went wrong - N2"+e); }            
        //}
        private void GetProjectsAndCollections()
        {
            //TeamProjectArrayList = new ArrayList();
            try
            {
                String ProQuery = "select p.ParentNodeSK as CollectionId,p.ProjectNodeSK as ProjectId,c.ProjectNodeName as CollectionName,p.ProjectNodeName as ProjectName from dbo.DimTeamProject p,Tfs_Warehouse.dbo.DimTeamProject c where p.ParentNodeSK=c.ProjectNodeSK and p.Depth=1";
                SqlConnection cn = new SqlConnection(ConfigurationManager.AppSettings["TfsDataSrc"]);
                cn.Open();
                SqlCommand cmd = new SqlCommand(ProQuery, cn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds, "Projects");
                foreach (DataRow row in ds.Tables["Projects"].Rows)
                {
                    String CollectionName = row["CollectionName"].ToString();
                    String ProjectName = row["ProjectName"].ToString();
                    int CollectionId = Convert.ToInt32(row["CollectionId"]);
                    int ProjectId = Convert.ToInt32(row["ProjectId"]);
                    //TeamProject TP = new TeamProject(CollectionId,CollectionName,ProjectName,ProjectId);                    
                    //TeamProjectArrayList.Add(TP);
                    //?
                    String AddProColQuery = "insert into dbo.TfsProjects(CollectionId,CollectionName,ProjectName,ProjectId) values(" + CollectionId + ",'" + CollectionName + "','" + ProjectName + "'," + ProjectId +")";
                    cn = new SqlConnection(ConfigurationManager.AppSettings["CustDataSrc"]);
                    cn.Open();
                    cmd = new SqlCommand(AddProColQuery, cn);
                    try
                    {
                        int x = cmd.ExecuteNonQuery();
                        ServiceEventLog("Inserting " + CollectionName + "-" + ProjectName);
                    }
                    catch (InvalidOperationException) { ServiceEventLog("Error inserting " + CollectionName + "-" + ProjectName); }
                    catch (SqlException) { ServiceEventLog(CollectionName + "-" + ProjectName + " already exists"); }
                }                
            }
            catch (ConfigurationErrorsException) { ServiceEventLog("No valid configuration file found - N3"); }
            catch (SqlException e) { ServiceEventLog("Database access error - N3"+e); }
            catch (Exception e){ ServiceEventLog("Something went wrong - N3"+ e); }
        }
        //private void SyncHistory()
        //{
        //    for (int i = 0; i < CollectionAL.Count; i++)
        //    {
        //        String HQuery = "select Title,[Changed Date],Rev,State,Reason,ID,"+
        //            "(select distinct(NamePart) from Tfs_" + CollectionAL[i] + ".dbo.Constants,Tfs_" + CollectionAL[i] + ".dbo.WorkItemsAre where ConstID=a.[Changed By]) as [Changed By],"+
        //            "(select distinct(NamePart) from Tfs_" + CollectionAL[i] + ".dbo.Constants,Tfs_" + CollectionAL[i] + ".dbo.WorkItemsAre where ConstID=a.[Assigned To]) as [Assigned To] " +
        //                    "FROM Tfs_" + CollectionAL[i] + ".dbo.WorkItemsWere a";
        //        try
        //        {
        //            SqlConnection cn = new SqlConnection(ConfigurationManager.AppSettings["TfsDataSrc"]);
        //            cn.Open();
        //            SqlCommand cmd = new SqlCommand(HQuery, cn);
        //            SqlDataAdapter da = new SqlDataAdapter(cmd);
        //            DataSet ds = new DataSet();
        //            da.Fill(ds, "WorkitemHistory");
        //            for (int j = 0; j < ds.Tables["WorkitemHistory"].Rows.Count; j++)
        //            {
        //                DataRow dr = ds.Tables["WorkitemHistory"].Rows[j];
        //                string AddHistoryQuery = "insert into Mydb.dbo.TfsWorkitemHistory(Title,State,Reason,[Assigned To],[Changed Date],[Changed By],Version) " + 
        //                        "values('"+dr["Title"]+"','"+dr["State"]+"','"+dr["Reason"]+"','"+dr["Assigned To"]+"','"+dr["Changed Date"]+"','"+dr["Changed By"]+"',"+dr["Rev"]+")";
        //                cn = new SqlConnection(ConfigurationManager.AppSettings["CustDataSrc"]);
        //                cn.Open();
        //                cmd = new SqlCommand(AddHistoryQuery, cn);
        //                try
        //                {
        //                    int x = cmd.ExecuteNonQuery();
        //                    ServiceEventLog("inserting history - " + dr["Title"]);
        //                }
        //                catch (InvalidOperationException) { ServiceEventLog("Error inserting " + dr["Title"]); }
        //                catch (SqlException) { ServiceEventLog(dr["Title"] + " already exists"); }
        //                cn.Close();
        //            }
        //            cn.Close();
        //        }
        //        catch (ConfigurationErrorsException) { ServiceEventLog("No valid configuration file found - N6"); }
        //        catch (SqlException e) { ServiceEventLog("Database access error - N6" + e); }
        //        catch (Exception e) { ServiceEventLog("Something went wrong - N6" + e); }
        //    }
        //}
        private void SyncWorkitems()
        {
            String SQuery = "select ID,Title,[Type],[Desc],CollectionName,ProjectName from dbo.TfsProjects,dbo.TfsWorkitems where Project=ProjectId and Sync=0";
            try
            {
                SqlConnection cn = new SqlConnection(ConfigurationManager.AppSettings["CustDataSrc"]);
                cn.Open();
                SqlCommand cmd = new SqlCommand(SQuery, cn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds, "Workitems");
                foreach (DataRow row in ds.Tables["Workitems"].Rows)
                {
                    String uri = ConfigurationManager.AppSettings["TfsUri"] + row["CollectionName"];
                    Uri collectionUri = new Uri(uri);
                    NetworkCredential myNetCredentials = new NetworkCredential(ConfigurationManager.AppSettings["TfsUsername"], ConfigurationManager.AppSettings["TfsPassword"]);
                    ICredentials myCredentials = (ICredentials)myNetCredentials;

                    TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(collectionUri, myCredentials);
                    WorkItemStore workItemStore = tpc.GetService<WorkItemStore>();
                    Project teamProject = workItemStore.Projects[row["ProjectName"].ToString()];
                    WorkItemType workItemType = teamProject.WorkItemTypes[row["Type"].ToString()];

                    WorkItem workItem = new WorkItem(workItemType);
                    workItem.Title = row["Title"].ToString();
                    workItem.Description = row["Desc"].ToString();
                    workItem.Save();
                    cmd = new SqlCommand("update dbo.TfsWorkitems set Sync=1,TfsId=" + workItem.Id + " where ID=" + row["ID"], cn);
                    try
                    {
                        int x = cmd.ExecuteNonQuery();
                        ServiceEventLog("Status - " + row["Title"] + " synced");
                    }
                    catch (InvalidOperationException) { ServiceEventLog("Error updating sync status"); }
                }
            }
            catch (ConfigurationErrorsException) { ServiceEventLog("No valid configuration file found - N4"); }
            catch (SqlException e) { ServiceEventLog("Database access error - N4"+e); }
            catch (Exception e){ ServiceEventLog("Something went wrong - N4"+e); }
        }
        private void GetWorkitems()
        {
            try
            {
                    //String AddWIQ = "select State,Reason,[Assigned To],Title,[Work Item Type],Words " +
                    //    "from Tfs_" + Collection + ".dbo.WorkItemsAre a,Tfs_" + Collection + ".dbo.xxTree b,Tfs_" + Collection + ".dbo.WorkItemLongTexts c " +
                    //    "where a.AreaID=b.ID and a.ID=c.ID and b.[Node Name]='" + Project + "'";
                    //String AddWIQ = "select State,Reason,[Assigned To],Title,[Work Item Type],Words "+
                    //        "from  Tfs_"+Collection+".dbo.xxTree b right join Tfs_"+Collection+".dbo.WorkItemsAre a left join Tfs_"+Collection+".dbo.WorkItemLongTexts c "+
                    //        "on a.ID=c.ID on a.AreaID=b.ID where b.[Node Name]='"+Project+"'";

                    //String AddWIQ = "select [Changed Date],Rev,State,Reason,NamePart as [Assigned To],[Created By],[Created Date],[Changed By],Title,[Work Item Type],Words " +
                    //        "from "+
                    //        "(select a.[Changed Date],a.Rev,State,Reason,[Assigned To],[Created By],[Created Date],[Changed By],Title,[Work Item Type],Words "+
                    //        "from  Tfs_"+Collection+".dbo.xxTree b right join Tfs_"+Collection+".dbo.WorkItemsAre a left join Tfs_"+Collection+".dbo.WorkItemLongTexts c "+
                    //        "on a.ID=c.ID on a.AreaID=b.ID where b.[Node Name]='"+Project+"') d,Tfs_"+Collection+".dbo.Constants e "+
                    //        "where d.[Assigned To]=e.ConstID";
                    //String AddWIQ = "select a.ID,a.[Changed Date],a.Rev,State,Reason,"+
                    //        "(select distinct(NamePart) from Tfs_" + CollectionName + ".dbo.Constants,Tfs_" + CollectionName + ".dbo.WorkItemsAre where ConstID=a.[Assigned To]) as [Assigned To]," +
                    //        "(select distinct(NamePart) from Tfs_" + CollectionName + ".dbo.Constants,Tfs_" + CollectionName + ".dbo.WorkItemsAre where ConstID=a.[Created By]) as [Created By],[Created Date]," +
                    //        "(select distinct(NamePart) from Tfs_" + CollectionName + ".dbo.Constants,Tfs_" + CollectionName + ".dbo.WorkItemsAre where ConstID=a.[Changed By]) as [Changed By],Title,[Work Item Type],Words " +
                    //        "from  Tfs_" + CollectionName + ".dbo.xxTree b right join Tfs_" + CollectionName + ".dbo.WorkItemsAre a left join Tfs_" + CollectionName + ".dbo.WorkItemLongTexts c " +
                    //        "on a.ID=c.ID on a.AreaID=b.ID where b.[Node Name]='"+ProjectName+"'";
                    String AddWIQ = "SELECT [System_Id] as ID"+
                                      ",[TeamProjectCollectionSK] as [Collection]"+
                                      ",[TeamProjectSK] as [Project]"+
                                      ",(select [Name] from [dbo].[DimPerson] where [PersonSK]=[System_AssignedTo__PersonSK]) as [Assigned To]"+
                                      ",(select [Name] from [dbo].[DimPerson] where [PersonSK]=[System_ChangedBy__PersonSK]) as [Changed By]"+
                                      ",(select [Name] from [dbo].[DimPerson] where [PersonSK]=[System_CreatedBy__PersonSK]) as [Created By]"+
                                      ",[System_WorkItemType] as [Work Item Type]"+
                                      ",[System_Title] as Title"+
                                      ",[System_ChangedDate] as [Changed Date]"+
                                      ",[System_State] as State"+
                                      ",[System_Rev] as [Rev]"+
                                      ",[System_Reason] as Reason"+
                                      ",[System_CreatedDate] as [Created Date]"+
                                      "FROM [dbo].[DimWorkItem]"+
                                      "WHERE [System_Rev]=1";

                    SqlConnection cn = new SqlConnection(ConfigurationManager.AppSettings["TfsDataSrc"]);
                    cn.Open();
                    SqlCommand cmd = new SqlCommand(AddWIQ, cn);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "Workitems");
                    foreach (DataRow row in ds.Tables["Workitems"].Rows)
                    {
                        String AddWI = "insert into dbo.TfsWorkitems(Project,Type,Title,[Assigned To],[Created By],[Created Date],[Changed By],[Changed Date],State,Reason,Version,Sync,TfsId) " +
                            "values(" + row["Project"] + ",'" + row["Work Item Type"] + "','" + row["Title"] + "','" +
                            row["Assigned To"] + "','" + row["Created By"] + "','" + DateTime.Parse(row["Created Date"].ToString()).ToLongDateString() + "','" + row["Changed By"] + "','" + DateTime.Parse(row["Changed Date"].ToString()).ToLongDateString() + "','" + 
                            row["State"] + "','" + row["Reason"] + "'," + row["Rev"] + ",1,"+row["ID"]+")";
                        cn = new SqlConnection(ConfigurationManager.AppSettings["CustDataSrc"]);
                        cn.Open();
                        cmd = new SqlCommand(AddWI, cn);
                        try
                        {
                            int x = cmd.ExecuteNonQuery();
                            ServiceEventLog("Inserting " + row["Title"]);
                        }
                        catch(InvalidOperationException)
                        {
                            ServiceEventLog("Error inserting " + row["Title"]);
                            String UpdateWI = "update dbo.TfsWorkitems " +
                                "set State='" + row["State"] + "',Reason='" + row["Reason"] + "',[Assigned To]='" + row["Assigned To"] + "',"+
                                "[Changed By]='" + row["Changed By"] + "',[Changed Date]='" + DateTime.Parse(row["Changed Date"].ToString()).ToLongDateString() + "',Version=" + row["Rev"] +
                                 " where TdsId=" + row["ID"];
                            cmd = new SqlCommand(UpdateWI, cn);
                            try
                            {
                                int x = cmd.ExecuteNonQuery();
                                ServiceEventLog("Updating " + row["Title"]);
                            }
                            catch (InvalidOperationException) { ServiceEventLog("Error updating " + row["Title"]); }
                            catch (SqlException) { ServiceEventLog("Error updating " + row["Title"]); }
                        }
                        catch (SqlException)
                        {
                            ServiceEventLog("Error inserting " + row["Title"]);
                            String UpdateWI = "update dbo.TfsWorkitems " +
                                "set State='" + row["State"] + "',Reason='" + row["Reason"] + "',[Assigned To]='" + row["Assigned To"] + "'," +
                                "[Changed By]='" + row["Changed By"] + "',[Changed Date]='" + DateTime.Parse(row["Changed Date"].ToString()).ToLongDateString() + "',Version=" + row["Rev"] +
                                " where TdsId=" + row["ID"];
                            cmd = new SqlCommand(UpdateWI, cn);
                            try
                            {
                                int x = cmd.ExecuteNonQuery();
                                ServiceEventLog("Updating " + row["Title"]);
                            }
                            catch (InvalidOperationException) { ServiceEventLog("Error updating " + row["Title"]); }
                            catch (SqlException) { ServiceEventLog("Error updating " + row["Title"]); }
                        }
                    }
            }
            catch (ConfigurationErrorsException) { ServiceEventLog("No valid configuration file found - N5"); }
            catch (SqlException e) { ServiceEventLog("Database access error - N5"+e); }
            catch { ServiceEventLog("Something went wrong - N5"); }
        }
        protected override void OnStop()
        {
            timer.Stop();
            timer.Dispose();
            ServiceEventLog("Stopping service");
            ServiceEventLog("------------------------------------------------------------------");
        }
        private void ServiceEventLog(String log)
        {
            try
            {
                FileStream fs = new FileStream("TfsSyncServiceLog.txt", FileMode.OpenOrCreate, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs);
                sw.BaseStream.Seek(0, SeekOrigin.End);
                sw.WriteLine(DateTime.Now + " : "+ log +".");
                sw.Flush();
                sw.Close();
            }
            catch { }
        }
    }
}
