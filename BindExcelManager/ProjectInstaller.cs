using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Threading.Tasks;
using System.ServiceProcess;
using System.Diagnostics;
namespace BindExcelManager
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();

            bindExcelInstaller = new ServiceInstaller();
            bindExcelInstaller.ServiceName = "BindExcelManager";
            bindExcelInstaller.DisplayName = "BindExcelManager";
            bindExcelInstaller.StartType = ServiceStartMode.Automatic;
            bindExcelInstaller.Description = "Service that manages excel file generation for BIND application";
            // kill the default event log installer
            bindExcelInstaller.Installers.Clear();
            EventLogInstaller logInstaller = new EventLogInstaller();

            logInstaller.Source = "BindExcelManager";
            logInstaller.Log = "ExcelLog";

            // Add the event log installer
            bindExcelInstaller.Installers.Add(logInstaller);
        }

        private void bindExcelInstaller_AfterInstall(object sender, InstallEventArgs e)
        {

        }
    }
}
