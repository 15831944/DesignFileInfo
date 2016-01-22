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
using System.Collections.Generic;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Forms;
using System.IO;
using DesignFile.Info;


namespace DesignFile.RevitPlugin
{
    [Transaction(TransactionMode.Manual)] 
    [Regeneration(RegenerationOption.Manual)]
    public class Commands : IExternalCommand

    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                
                FolderBrowserDialog fbgRootDir;
                SaveFileDialog fdCSVPath;
                DirectoryInfo rootDirInfo;

            
                fbgRootDir = new System.Windows.Forms.FolderBrowserDialog();
                // Set the help text description for the FolderBrowserDialog. 
                fbgRootDir.Description =
                    "Select the root folder that contains RVT files and/or subfolders of RVT files.";
                // Do not allow the user to create new files via the FolderBrowserDialog. 
                fbgRootDir.ShowNewFolderButton = false;
                // Default to the My Computer folder. 
                fbgRootDir.RootFolder = Environment.SpecialFolder.MyComputer;

            
                fdCSVPath = new System.Windows.Forms.SaveFileDialog();
                fdCSVPath.DefaultExt = "csv";
                fdCSVPath.Filter = "Comma-separated files (*.csv)|*.csv";
                fdCSVPath.AddExtension = true;
                fdCSVPath.Title = "Select the location of the results as a CSV file.";

                DialogResult result = fbgRootDir.ShowDialog();
                if (result == DialogResult.OK)
                {
                    fdCSVPath.InitialDirectory = fbgRootDir.SelectedPath;
                    if (fdCSVPath.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            FileInfo fi = new FileInfo(fdCSVPath.FileName);
                            if (fi.Exists) {
                                try
                                {
                                    fi.Delete();                                     
                                }
                                catch (IOException ex)
                                {
                                    
                                    DialogResult dr = MessageBox.Show(fi.FullName + " is in use or read-only.  If you have the file open, close it, then press Retry.","File in Use",
                                          MessageBoxButtons.RetryCancel,
                                         MessageBoxIcon.Error);
                                    if (dr == DialogResult.Cancel)
                                    {
                                        return Result.Cancelled;
                                    }
                                    else
                                    {
                                        try
                                        {
                                            fi.Delete();
                                        }
                                        catch (Exception)
                                        {
                                            message = "Unable to overwrite the existing CSV file.  Make sure the file you selected is closed or deleted, and try again.";
                                            return Result.Failed;
                                        }

                                    }
                                } 
                            }
                            rootDirInfo = new DirectoryInfo(fbgRootDir.SelectedPath);

                            using (System.IO.StreamWriter file = new System.IO.StreamWriter(fdCSVPath.FileName, true))
                            {
                                file.WriteLine(DateTime.Now);
                                file.WriteLine("\"File Name\",\"Version\",\"Is Central\",\"Central File Name\",\"Link Type\",\"Path\",\"Absolute Path\"");
                            }

                            ProcessFolder(rootDirInfo, fdCSVPath.FileName);


                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace, "Error getting file information",
                                         //MessageBoxButtons.OK,
                                         //MessageBoxIcon.Error);
                            throw new Exception("Error getting file information", ex);
                        }
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

        }


        static void ProcessFolder (DirectoryInfo directory, string logFilePath)
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
                try
                {
                    file.WriteLine(revitFileInfo.ToCSV());
                }
                catch (Exception ex)
                {
                    WriteError("Error writing to the CSV file: " + logFilePath + " - " + ex.Message, logFilePath);
                }
                
            }
        }

        //static void WriteAutoCADInfo(AutoCADFileInfo autoCADFileInfo, string logFilePath)
        //{
        //    using (System.IO.StreamWriter file = new System.IO.StreamWriter(logFilePath, true))
        //    {
        //        file.WriteLine(autoCADFileInfo.ToCSV());
        //    }
        //}

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




    
    
    [Transaction(TransactionMode.Manual)] 
    [Regeneration(RegenerationOption.Manual)]
    public class UI : IExternalApplication
    {

        public Result OnStartup(UIControlledApplication application) {

 	        //Set up the button in the ribbon
            String path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            String caption = "Linked\nFile\nInformation";
            PushButtonData d = new PushButtonData(caption, caption, path, "DesignFile.RevitPlugin.Commands");
            d.AvailabilityClassName = "DesignFile.RevitPlugin.Availability";

            //Add Panel and Button
            RibbonPanel p = application.CreateRibbonPanel("Linked File Information");
            PushButton b = p.AddItem(d) as PushButton;
            b.ToolTip = "Opens a folder of Revit files and lists the linked reference files.";

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application) {
 	        return Result.Succeeded;
        }

    }

    public class Availability : IExternalCommandAvailability
    {
	    public bool IsCommandAvailable1(UIApplication applicationData, CategorySet selectedCategories)
	    {
		    return true;
	    }
	    bool IExternalCommandAvailability.IsCommandAvailable(UIApplication applicationData, CategorySet selectedCategories)
	    {
		    return IsCommandAvailable1(applicationData, selectedCategories);
	    }
    }

    public class RevitFile
    {

        public static RevitFileInfo GetRevitFileInfo(string path)
        {
            try
            {
                RevitFileInfo rfi = new RevitFileInfo();
                rfi.FileName = path;

                if (!DesignFile.Info.BasicFileInfo.StructuredStorageUtils.IsFileStucturedStorage(path))
                {
                    throw new NotSupportedException(
                        "File is not a structured storage file");
                }

                var rawData = DesignFile.Info.BasicFileInfo.GetRawBasicFileInfo(path);

                var rawString = System.Text.Encoding.Unicode
                .GetString(rawData);

                var fileInfoData = rawString.Split(
                    new string[] { "\0", "\r\n" },
                    StringSplitOptions.RemoveEmptyEntries);

                foreach (var info in fileInfoData)
                {
                    rfi.RawData += info;
                    if (info.Contains("sharing:"))
                    {
                        int i = info.IndexOf("sharing: ") + 9;
                        rfi.IsCentral = (info.Substring(i, info.Length - i) == "Central");
                    };
                    if (info.Contains("Central Model Path"))
                    {
                        rfi.CentralFileName = info.Substring(19);
                    };
                    if (info.Contains("Revit Build:"))
                    {
                        rfi.Version = info.Substring(13);
                    };
                }

                rawString = System.Text.Encoding.BigEndianUnicode
                .GetString(rawData);

                fileInfoData = rawString.Split(
                new string[] { "\0", "\r\n" },
                StringSplitOptions.RemoveEmptyEntries);

                foreach (var info in fileInfoData)
                {
                    rfi.RawData += info;
                    if (info.Contains("sharing:"))
                    {
                        int i = info.IndexOf("sharing: ") + 9;
                        rfi.IsCentral = (info.Substring(i, info.Length - i) == "Central");
                    };
                    if (info.Contains("Central Model Path"))
                    {
                        rfi.CentralFileName = info.Substring(19);
                    };
                    if (info.Contains("Revit Build:"))
                    {
                        rfi.Version = info.Substring(13);
                    };

                }

                rfi.Links = GetLinks(rfi.FileName);

                return rfi;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }

        private static List<RevitLink> GetLinks(string location)
        {
            List<RevitLink> links = new List<RevitLink>();
            try
            {
                ModelPath path = ModelPathUtils.ConvertUserVisiblePathToModelPath(location);

                TransmissionData transData = TransmissionData.ReadTransmissionData(path);
                ICollection<ElementId> externalReferences = default(ICollection<ElementId>);

                if (transData != null)
                {
                    // collect all (immediate) external references in the model

                    externalReferences = transData.GetAllExternalFileReferenceIds();

                    if (externalReferences.Count > 0)
                    {
                        foreach (ElementId refId in externalReferences)
                        {
                            ExternalFileReference extRef = transData.GetLastSavedReferenceData(refId);
                            if (extRef.IsValidObject)
                            {
                                if (extRef.ExternalFileReferenceType == ExternalFileReferenceType.CADLink | extRef.ExternalFileReferenceType == ExternalFileReferenceType.DWFMarkup | extRef.ExternalFileReferenceType == ExternalFileReferenceType.RevitLink)
                                {
                                    RevitLink rl = new RevitLink();
                                    rl.LinkType = extRef.ExternalFileReferenceType.ToString();
                                    rl.AbsolutePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(extRef.GetAbsolutePath());
                                    rl.Path = ModelPathUtils.ConvertModelPathToUserVisiblePath(extRef.GetPath());
                                    links.Add(rl);
                                }
                            }
                        }

                    }
                }
            }
            catch (Exception)
            {

            }
            return links;
        }


    }

    public class RevitLink
    {
        public string AbsolutePath { get; set; }
        public string Path { get; set; }
        public string LinkType { get; set; }
    }


    public class RevitFileInfo
    {
        public bool IsCentral { get; set; }
        public string FileName { get; set; }
        public string CentralFileName { get; set; }
        public string Version { get; set; }
        public string RawData { get; set; }
        public List<RevitLink> Links { get; set; }

        public string ToString()
        {
            string ret = string.Empty;

            ret += "File Name: " + FileName;
            ret += "Version: " + Version;
            ret += "Is Central File: " + Convert.ToString(IsCentral);
            if (IsCentral)
            {
                ret += "Central File Name: " + CentralFileName;
            }

            return ret;
        }

        public string ToCSV()
        {
            string ret = string.Empty;

            if (Links.Count == 0)
            {
                ret += "\"" + FileName + "\",";
                ret += "\"" + Version + "\",";
                ret += "\"" + Convert.ToString(IsCentral) + "\",";
                ret += "\"" + CentralFileName + "\",";
            }
            else
            {
                int i = 0;
                foreach (RevitLink link in Links)
                {
                    i += 1;
                    ret += "\"" + FileName + "\",";
                    ret += "\"" + Version + "\",";
                    ret += "\"" + Convert.ToString(IsCentral) + "\",";
                    ret += "\"" + CentralFileName + "\",";
                    ret += "\"" + link.LinkType + "\",";
                    ret += "\"" + link.Path + "\",";
                    ret += "\"" + link.AbsolutePath + "\",";
                    if (i!=Links.Count) {
                        ret += System.Environment.NewLine;
                    }
                }
            }

            return ret;
        }
    }


}
