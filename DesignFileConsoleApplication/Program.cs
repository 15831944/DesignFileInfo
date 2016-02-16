// (C) Copyright 2014 by Autodesk, Inc. 
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted, 
// provided that the above copyright notice appears in all copies and 
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting 
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC. 
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to 
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//


using System;
using System.IO;
//using System.IO.Packaging;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DesignFile.Info;

namespace DesignFileConsoleApplication
{
    class Program
    {

        [STAThreadAttribute]
        static void Main()
        {

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            FolderBrowserDialog fbgRootDir;
            SaveFileDialog fdCSVPath;
            DirectoryInfo rootDirInfo;


            fbgRootDir = new System.Windows.Forms.FolderBrowserDialog();
            // Set the help text description for the FolderBrowserDialog. 
            fbgRootDir.Description =
                "Select the root folder that you want to traverse.";
            // Do not allow the user to create new files via the FolderBrowserDialog. 
            fbgRootDir.ShowNewFolderButton = false;
            // Default to the My Computer folder. 
            fbgRootDir.RootFolder = Environment.SpecialFolder.MyComputer;


            fdCSVPath = new System.Windows.Forms.SaveFileDialog();
            fdCSVPath.DefaultExt = "csv";
            fdCSVPath.Filter = "Comma-separated files (*.csv)|*.csv";
            fdCSVPath.AddExtension = true;
            fdCSVPath.Title = "Select the location of the output logs";

            DialogResult result = fbgRootDir.ShowDialog();
            if (result == DialogResult.OK)
            {
                fdCSVPath.InitialDirectory = fbgRootDir.SelectedPath;
                if (fdCSVPath.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (File.Exists(fdCSVPath.FileName)) { File.Delete(fdCSVPath.FileName); }
                        rootDirInfo = new DirectoryInfo(fbgRootDir.SelectedPath);

                        using (System.IO.StreamWriter file = new System.IO.StreamWriter(fdCSVPath.FileName, true))
                        {
                            file.WriteLine(DateTime.Now);
                            //file.WriteLine("\"File Name\",\"Version\",\"Is Central\",\"Central File Name\",\"Link Type\",\"Path\",\"Absolute Path\"");
                            file.WriteLine("\"File Name\",\"Version\",\"Is Central\",\"Central File Name\"");
                        }

                        //using (System.IO.StreamWriter file = new System.IO.StreamWriter(rootDirInfo.FullName + "\\DesignFiles_AutoCADFiles.csv", true))
                        //{
                        //    file.WriteLine(DateTime.Now);
                        //    file.WriteLine("\"File Name\",\"Version\"");
                        //}

                        ProcessFolder(rootDirInfo, fdCSVPath.FileName);


                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace, "Error getting file information",
                                     MessageBoxButtons.OK,
                                     MessageBoxIcon.Error);
                    }
                }
            }
        }

        static void ProcessFolder(DirectoryInfo directory, string logFilePath)
        {
            // Process the Revit files first
            var rvtFiles = directory.EnumerateFiles("*.rvt");
            foreach (FileInfo file in rvtFiles)
            {
                RevitFileInfo rfi;
                try
                {
                    Console.WriteLine(file.FullName);
                    rfi = RevitFile.GetRevitFileInfo(file.FullName);
                    if (rfi != null)
                    {
                        try
                        {
                            WriteRevitInfo(rfi, logFilePath);
                        }
                        catch (Exception ex)
                        {
                            WriteError("Error writing the file information for: " + file.FullName + " - " + ex.Message, logFilePath);
                            continue;
                        }
                    }
                }
                catch (PathTooLongException ex)
                {
                    WriteError("The path or name of a file is too long (260+ characters) in: " + file.Directory + " - " + ex.Message, logFilePath);
                    continue;
                }
                catch (Exception ex)
                {
                    WriteError("Error getting Revit File Info for: " + file.FullName + " - " + ex.Message, logFilePath);
                    continue;
                }


            }

            //// Process the AutoCAD files next
            //var dwgFiles = directory.EnumerateFiles("*.dwg");
            //foreach (FileInfo file in dwgFiles)
            //{
            //    AutoCADFileInfo afi;
            //    try
            //    {
            //        afi = AutoCADFile.GetAutoCADFileInfo(file.FullName);
            //        if (afi != null)
            //        {
            //            try
            //            {
            //                WriteAutoCADInfo(afi, logFilePath);
            //            }
            //            catch (Exception ex)
            //            {
            //                WriteError("Error writing the file information for: " + file.FullName + " - " + ex.Message, logFilePath);
            //            }
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        WriteError("Error getting Revit File Info for: " + file.FullName + " - " + ex.Message, logFilePath);
            //    }


            //}

            // Process the folders next
            var subDirectories = directory.EnumerateDirectories();
            foreach (DirectoryInfo subdirectory in subDirectories)
            {
                Console.WriteLine("_____________________");
                Console.WriteLine(subdirectory.FullName);
                try
                {
                    ProcessFolder(subdirectory, logFilePath);
                }
                catch (PathTooLongException ex)
                {
                    WriteError("The folder path or name of a file is too long (260+ characters) in: " + subdirectory.Parent.FullName + " - " + ex.Message, logFilePath);
                    continue;
                }
                catch (Exception ex)
                {
                    WriteError("Error traversing subfolders of:  " + subdirectory.FullName + " - " + ex.Message, logFilePath);
                    continue;
                }

            }
        }

        static void WriteRevitInfo(RevitFileInfo revitFileInfo, string logFilePath)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(logFilePath, true))
            {
                file.WriteLine(revitFileInfo.ToCSV());
            }
        }

        static void WriteAutoCADInfo(AutoCADFileInfo autoCADFileInfo, string logFilePath)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(logFilePath, true))
            {
                file.WriteLine(autoCADFileInfo.ToCSV());
            }
        }

        static void WriteError(string message, string logFilePath)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(logFilePath + "_Errors.txt", true))
            {
                file.WriteLine(message);
            }
        }

        static void TestWritingToFolder(DirectoryInfo directory)
        {
            try
            {
                System.IO.File.WriteAllText(directory.FullName + "\\TestFile.txt", "This is a test file");
                File.Delete(directory.FullName + "\\TestFile.txt");
            }
            catch (Exception ex)
            {
                throw new Exception("Error writing to the root folder you selected, make sure this folder is writable for the logs.", ex);
            }

        }

    }
}
