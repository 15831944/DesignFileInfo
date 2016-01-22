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
using System.IO.Packaging;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections.Generic;
//using Autodesk.Revit;
//using Autodesk.Revit.DB;

namespace DesignFile.Info
{
    public class RevitFile
    {

        public static RevitFileInfo GetRevitFileInfo(string path)
        {
            try
            {
                RevitFileInfo rfi = new RevitFileInfo();
                rfi.FileName = path;

                if (!BasicFileInfo.StructuredStorageUtils.IsFileStucturedStorage(path))
                {
                    throw new NotSupportedException(
                        "File is not a structured storage file");
                }

                var rawData = BasicFileInfo.GetRawBasicFileInfo(path);

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

                //rfi.Links = GetLinks(rfi.FileName);

                return rfi;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }

        //private static List<RevitLink> GetLinks(string location)
        //{
        //    List<RevitLink> links = new List<RevitLink>();
        //    try
        //    {
        //        ModelPath path = ModelPathUtils.ConvertUserVisiblePathToModelPath(location);
                
        //        TransmissionData transData = TransmissionData.ReadTransmissionData(path);
        //        ICollection<ElementId> externalReferences = default(ICollection<ElementId>);

        //        if (transData != null)
        //        {
        //            // collect all (immediate) external references in the model

        //            externalReferences = transData.GetAllExternalFileReferenceIds();

        //            if (externalReferences.Count > 0)
        //            {
        //                foreach (ElementId refId in externalReferences)
        //                {
        //                    ExternalFileReference extRef = transData.GetLastSavedReferenceData(refId);
        //                    if (extRef.IsValidObject)
        //                    {
        //                        if (extRef.ExternalFileReferenceType == ExternalFileReferenceType.CADLink | extRef.ExternalFileReferenceType == ExternalFileReferenceType.DWFMarkup | extRef.ExternalFileReferenceType == ExternalFileReferenceType.RevitLink)
        //                        {
        //                            RevitLink rl = new RevitLink();
        //                            rl.LinkType = extRef.ExternalFileReferenceType.ToString();
        //                            rl.AbsolutePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(extRef.GetAbsolutePath());
        //                            rl.Path = ModelPathUtils.ConvertModelPathToUserVisiblePath(extRef.GetPath());
        //                        }
        //                    }
        //                }

        //            }
        //        }
        //    }
        //    catch (Exception)
        //    {

        //    }
        //    return links;
        //}


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

            //if (Links.Count == 0)
            //{
                ret += "\"" + FileName + "\",";
                ret += "\"" + Version + "\",";
                ret += "\"" + Convert.ToString(IsCentral) + "\",";
                ret += "\"" + CentralFileName + "\",";
            //}
            //else
            //{
            //    foreach (RevitLink link in Links)
            //    {
            //        ret += "\"" + FileName + "\",";
            //        ret += "\"" + Version + "\",";
            //        ret += "\"" + Convert.ToString(IsCentral) + "\",";
            //        ret += "\"" + CentralFileName + "\",";
            //        ret += "\"" + link.LinkType + "\",";
            //        ret += "\"" + link.Path + "\",";
            //        ret += "\"" + link.AbsolutePath + "\",";
            //    }
            //}

            return ret;
        }
    }




}



