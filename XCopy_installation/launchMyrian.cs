/*
 * Created by Ranorex
 * User: hamza.erraji
 * Date: 29/03/2016
 * Time: 10:34
 * 
 * 
 */
using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace launchMyrian
{
    /// <summary>
    /// Description de LaunchMyrian.
    /// 
    /// Ce script permet de tester le lancement de Myrian avec les Executables sans utiliser le Setup. Pour ce fait, Les Executables de la derniere version ainsi que les dépendances fonctionnelles
    /// vont être copiés d'un dossier situé dans le "DépotCompilationAutomatique" pour être déposer dans un dossier dont le chemin va être saisie en 
    /// argument de la ligne de commande. Le nom du dossier de dépôt est "Myrian_mainline_bXXXXXXXX" (Avec XXXXXX le numero de build et la date de création).
    /// 
    /// Exemple de ligne de commande : launchMyrian.exe /tcpa:TestExec:codeline=mainLine /tcpa:TestExec:destDir=C:\Users\Desktop
    /// Les arguments de cette ligne de commandes sont :
    /// - codeLine : "mainline" ou "prep" (Pour l'instant le script ne gère pas la codeline "prep") 
    /// - destDir : Chemin du fichier de dépôt
    /// Le "TestCase" dans ce cas est representé par "TestExec" 
    /// </summary>
    [TestModule("B70E9F2E-2070-4F19-B3CA-39A37AD43BF0", ModuleType.UserCode, 1)]
    public class LaunchMyrian : ITestModule
    {
       
        public LaunchMyrian()
        {
            
        }

      	// CopyDir permet de copier un dossier avec son contenu de "sourceDir" à "desDir"
         public static void CopyDir(string sourceDir, string desDir)
       	 {
        	   DirectoryInfo dir = new DirectoryInfo(sourceDir);
		       if (dir.Exists)
		       {
			     string realDestDir;
			     if(dir.Root.Name != dir.Name && dir.Name != "exec_64bit" && dir.Name != "Myrian_Delta_Dependencies")
			     {
			       realDestDir = Path.Combine(desDir,dir.Name);
				   if(!Directory.Exists(realDestDir))
				   Directory.CreateDirectory(realDestDir);
			     }
			     else realDestDir = desDir;
			     foreach(string d in Directory.GetDirectories(sourceDir))
			      CopyDir(d, realDestDir);
			     foreach(string file in Directory.GetFiles(sourceDir))
			     {
			       string fileNamedest = Path.Combine(realDestDir,Path.GetFileName(file));
				   File.Copy(file,fileNamedest,true);
		         }
	           }
          }	
         
      	public static void modifyAttribute(string path)
        {

      		// Changer attribut "Lecture seule" pour les fichiers du dossier "racine"
      		DirectoryInfo dire = new DirectoryInfo(path);
			FileSystemInfo[] infos = dire.GetFileSystemInfos("*",SearchOption.AllDirectories);
			
			foreach (FileSystemInfo dri in infos)
            {
                if ((dri.Attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    dri.Attributes = FileAttributes.Normal;
                }
            }

        }
          
        
         // Déclaration d'une variable codeLine qui va être saisie en argument de la ligne de commande
         string _codeLine = "mainline";
         [TestVariable("27E7F754-E389-418D-BED8-1A20E9D1F08D")]
         public string codeLine
         {
         	get { return _codeLine; }
         	set { _codeLine = value; }
         }
         
         // Déclaration d'une variable destDir qui va être saisie en argument de la ligne de commande
         string _destDir = "C:" + "\\" + "Myrian";
         [TestVariable("75F04BAF-D99E-445C-898A-B74DEB162DBB")]
         public string destDir
         {
         	get { return _destDir; }
         	set { _destDir = value; }
         }
         
        void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.0;
            
            // Déclaration de variable contenant le chemin de la source des Executable et DLL
            string sourceDirExec = @"\\192.168.1.223\DepotCompilationAutomatique";
            string sourceDirDepend = @"\\192.168.1.223\DepotCompilationAutomatique\Myrian_Delta_Dependencies";
            // Le nom du fichier dans lequelle les Execs et DLL vont être copiés
            string folderMyrian;

            // Récuperation des dossiers dont le nom commence par "mainline ou 2.0_prep" 
	        DirectoryInfo dirInfo = new DirectoryInfo(sourceDirExec);
	        
	        // codeLine = mainline
	        if (String.Compare(codeLine,"mainline") == 0)
	        {
	       		 DirectoryInfo[] dirInfos = dirInfo.GetDirectories("mainline_b*");
	       		 // Recupération de la derniere version en prenant le dernier build
	       		 foreach(DirectoryInfo  d in dirInfos)
	       		 {
	        		if(String.Compare(d.Name,codeLine) > 0)
	         		{
	        			codeLine = d.Name;
	          		}
	         	 }
	         }
	        
	        // codeLine = 2.0_prep
	        
	        else if (String.Compare(codeLine,"2.0_prep") == 0)
	        {
	        	DirectoryInfo[] dirInfos = dirInfo.GetDirectories("2.0_prep_b*");
	       		 // Recupération de la derniere version en prenant le dernier build
	       		 foreach(DirectoryInfo  d in dirInfos)
	       		 {
	        		if(String.Compare(d.Name,codeLine) > 0)
	         		{
	        			codeLine = d.Name;
	          		}
	         	 }
	        }
	        
	        else
	        {
	        	Validate.IsFalse(true,"Veuillez saisir une codeLine valide :  'mainline ou 2.0_prep'");
	        }
	        	

            // Nom du fichier de dépôt
            folderMyrian = "Myrian"+"_"+codeLine;
            
            // Chemin du fichier de dépôt
            destDir = Path.Combine(destDir,folderMyrian);
            
            // Récuperation de chemin des Executables de la dernière version
            sourceDirExec = Path.Combine(sourceDirExec,codeLine);
            sourceDirExec = Path.Combine(sourceDirExec,"exec_64bit"); 
            
            // Affichage de quelques informations dans le rapport de Ranorex 
            Ranorex.Report.Info("Répertoire source des Executables:" + sourceDirExec);
            Ranorex.Report.Info("Répertoire source des dépendances fonctionnelles:" + sourceDirDepend);
            Ranorex.Report.Info("Répertoire de destination des Exec et DLL :" + destDir);  
            Ranorex.Report.Info("Copie en cours ...");
            
            // Copie des Executables 
		    CopyDir(sourceDirExec,destDir);
		    
		    // Modification de l'attribut "Lecture seul"
		    modifyAttribute(Path.Combine(destDir,"CRF"));
		   
          	// Copie des DLL
            CopyDir(sourceDirDepend,destDir); 
            
            // Lancer Myrian
            Host.Local.RunApplication(Path.Combine(destDir,"myrian.exe"), "", destDir, false);
       
		}
     }
}
