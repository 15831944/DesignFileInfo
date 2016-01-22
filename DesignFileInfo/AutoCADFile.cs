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

namespace DesignFile.Info

{
    public class AutoCADFile
    {

        public static AutoCADFileInfo GetAutoCADFileInfo(string path)
        {
            try
            {
                AutoCADFileInfo afi = new AutoCADFileInfo();
                afi.FileName = path;
                afi.Version = DwgVersion(path);
                return afi;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }

        private static string DwgVersion(string filename)
        {
            using (StreamReader reader = new StreamReader(filename))
            {
                switch (reader.ReadLine().Substring(0, 6))
                {
                    case "AC1027": return "AutoCAD 2013";
                    case "AC1024": return "AutoCAD 2010";
                    case "AC1021": return "AutoCAD 2007";
                    case "AC1018": return "AutoCAD 2004";
                    case "AC1015": return "AutoCAD 2000";
                    case "AC1014": return "AutoCAD R14";
                    default: return "Prior AutoCAD R14";
                }
            }
        }



    }



    public class AutoCADFileInfo
    {

        public string FileName { get; set; }
        public string Version { get; set; }
        public string RawData { get; set; }
        public string[] LinkedFiles { get; set; }
        public bool IsEMR { get; set; } 



        public string ToString()
        {
            string ret = string.Empty;

            ret += "File Name: " + FileName;
            ret += "Version: " + Version;


            return ret;
        }

        public string ToCSV()
        {
            string ret = string.Empty;

            ret += "\"" + FileName + "\",";
            ret += "\"" + Version + "\",";
            //ret += "\"" + RawData + "\",";

            return ret;
        }


    }

    public class LinkedFile
    {
        public string FileName { get; set; }
        public string Path { get; set; }
    }

}