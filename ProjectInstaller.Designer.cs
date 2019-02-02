namespace TfsSyncService
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tfsSyncServiceProcessInstaller = new System.ServiceProcess.ServiceProcessInstaller();
            this.tfsSyncServiceInstaller = new System.ServiceProcess.ServiceInstaller();
            // 
            // tfsSyncServiceProcessInstaller
            // 
            this.tfsSyncServiceProcessInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.tfsSyncServiceProcessInstaller.Password = null;
            this.tfsSyncServiceProcessInstaller.Username = null;
            // 
            // tfsSyncServiceInstaller
            // 
            this.tfsSyncServiceInstaller.DisplayName = "TfsSync";
            this.tfsSyncServiceInstaller.ServiceName = "SyncService";
            this.tfsSyncServiceInstaller.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.tfsSyncServiceProcessInstaller,
            this.tfsSyncServiceInstaller});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller tfsSyncServiceProcessInstaller;
        private System.ServiceProcess.ServiceInstaller tfsSyncServiceInstaller;
    }
}