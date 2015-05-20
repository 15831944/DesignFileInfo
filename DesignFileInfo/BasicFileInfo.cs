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
    public class BasicFileInfo
    {

        private const string StreamName = "BasicFileInfo";

        public static byte[] GetRawBasicFileInfo(string revitFileName)
        {
            if (!StructuredStorageUtils.IsFileStucturedStorage(
              revitFileName))
            {
                throw new NotSupportedException(
                  "File is not a structured storage file");
            }

            using (StructuredStorageRoot ssRoot =
                new StructuredStorageRoot(revitFileName))
            {
                if (!ssRoot.BaseRoot.StreamExists(StreamName))
                    throw new NotSupportedException(string.Format(
                      "File doesn't contain {0} stream", StreamName));

                StreamInfo imageStreamInfo =
                    ssRoot.BaseRoot.GetStreamInfo(StreamName);

                using (Stream stream = imageStreamInfo.GetStream(
                  FileMode.Open, FileAccess.Read))
                {
                    byte[] buffer = new byte[stream.Length];
                    stream.Read(buffer, 0, buffer.Length);
                    return buffer;
                }
            }
        }

        public static class StructuredStorageUtils
        {
            [DllImport("ole32.dll")]
            static extern int StgIsStorageFile(
              [MarshalAs(UnmanagedType.LPWStr)]
      string pwcsName);

            public static bool IsFileStucturedStorage(
              string fileName)
            {
                int res = StgIsStorageFile(fileName);

                if (res == 0)
                    return true;

                if (res == 1)
                    return false;

                throw new FileNotFoundException(
                  "File not found", fileName);
            }
        }

        public class StructuredStorageException : Exception
        {
            public StructuredStorageException()
            {
            }

            public StructuredStorageException(string message)
                : base(message)
            {
            }

            public StructuredStorageException(
              string message,
              Exception innerException)
                : base(message, innerException)
            {
            }
        }

        public class StructuredStorageRoot : IDisposable
        {
            StorageInfo _storageRoot;

            public StructuredStorageRoot(Stream stream)
            {
                try
                {
                    _storageRoot
                      = (StorageInfo)InvokeStorageRootMethod(
                        null, "CreateOnStream", stream);
                }
                catch (Exception ex)
                {
                    throw new StructuredStorageException(
                      "Cannot get StructuredStorageRoot", ex);
                }
            }

            public StructuredStorageRoot(string fileName)
            {
                try
                {
                    _storageRoot
                      = (StorageInfo)InvokeStorageRootMethod(
                        null, "Open", fileName, FileMode.Open,
                        FileAccess.Read, FileShare.Read);
                }
                catch (Exception ex)
                {
                    throw new StructuredStorageException(
                      "Cannot get StructuredStorageRoot", ex);
                }
            }

            private static object InvokeStorageRootMethod(
              StorageInfo storageRoot,
              string methodName,
              params object[] methodArgs)
            {
                Type storageRootType
                  = typeof(StorageInfo).Assembly.GetType(
                    "System.IO.Packaging.StorageRoot",
                    true, false);

                object result = storageRootType.InvokeMember(
                  methodName,
                  BindingFlags.Static | BindingFlags.Instance
                  | BindingFlags.Public | BindingFlags.NonPublic
                  | BindingFlags.InvokeMethod,
                  null, storageRoot, methodArgs);

                return result;
            }

            private void CloseStorageRoot()
            {
                InvokeStorageRootMethod(_storageRoot, "Close");
            }

            #region Implementation of IDisposable

            public void Dispose()
            {
                CloseStorageRoot();
            }

            #endregion

            public StorageInfo BaseRoot
            {
                get { return _storageRoot; }
            }

        }

    }

}