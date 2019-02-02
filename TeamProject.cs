using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TfsSyncService
{
    class TeamProject
    {
        public int CollectionId { get; set; }
        public string CollectionName { get; set; }
        public string ProjectName { get; set; }
        public int ProjectId { get; set; }
        public TeamProject(int CollectionId,string CollectionName,string ProjectName,int ProjectId)
        {
            this.CollectionId = CollectionId;
            this.CollectionName = CollectionName;
            this.ProjectName = ProjectName;
            this.ProjectId = ProjectId;
        }
    }
}
