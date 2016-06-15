/*
 * Created by Ranorex
 * User: hamza.erraji
 * Date: 14/06/2016
 * Time: 11:41
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using Ranorex.Core.Reporting;
using WinForms = System.Windows.Forms;
using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace launchMyrian
{
    /// <summary>
    /// Description of deleteOldReport.
    /// </summary>
    [TestModule("9D10432C-57D3-4E76-92EF-86099DD83B07", ModuleType.UserCode, 1)]
    public class deleteOldReport : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public deleteOldReport()
        {
            // Do not delete - a parameterless constructor is required!
        }

        /// <summary>
        /// Performs the playback of actions in this module.
        /// </summary>
        /// <remarks>You should not call this method directly, instead pass the module
        /// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
        /// that will in turn invoke this method.</remarks>
        void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.0;
            
            
           
           string pathReport = @"E:\";
           string nameTestSuite =TestSuite.Current.Name;
           
            pathReport = Path.Combine(pathReport,nameTestSuite);
            pathReport = Path.Combine(pathReport,nameTestSuite);
            pathReport = Path.Combine(pathReport,"bin");
            pathReport = Path.Combine(pathReport,"debug");
            pathReport = Path.Combine(pathReport,string.Concat(nameTestSuite,".rxzlog"));
            
            Ranorex.Report.Info(pathReport);
            
            if(System.IO.File.Exists(pathReport))
            {
               	System.IO.File.Delete(pathReport);
               	Ranorex.Report.Info("Ancien rapport supprimé");
            }
            else
               {
            	
               	Ranorex.Report.Info("Aucun rapport trouvé");	
               }
        }
    }
}
