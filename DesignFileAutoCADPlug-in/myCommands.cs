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
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(DesignFileAutoCADPlug_in.Commands))]

namespace DesignFileAutoCADPlug_in
{

    // This class is instantiated by AutoCAD for each document when
    // a command is called by the user the first time in the context
    // of a given document. In other words, non static data in this class
    // is implicitly per-document!
    public class Commands
    {
        // The CommandMethod attribute can be applied to any public  member 
        // function of any public class.
        // The function should take no arguments and return nothing.
        // If the method is an intance member then the enclosing class is 
        // intantiated for each document. If the member is a static member then
        // the enclosing class is NOT intantiated.
        //
        // NOTE: CommandMethod has overloads where you can provide helpid and
        // context menu.

        // Modal Command with localized name
        [CommandMethod("DesignFileGroup", "GetFileInfo", "GetFileInfo", CommandFlags.Modal)]
        public void GetDesignFileInfo() // This method can have any name
        {
            //// Put your command code here
            //Document doc = Application.DocumentManager.MdiActiveDocument;
            //Editor ed;
            //if (doc != null)
            //{
            //    ed = doc.Editor;
            //    ed.WriteMessage("Hello, this is your first command.");

            //}



        

        Database db = new Database(false, true);
        using (db)
        {
            try
            {
                db.ReadDwgFile(
                    res.StringResult,
                    System.IO.FileShare.Read,
                    false,
                    ""
                );
            }
            catch (System.Exception)
            {
                //ed.WriteMessage("\nUnable to read drawing file.");
            return;
          }

          Transaction tr = db.TransactionManager.StartTransaction();
          using (tr)
          {
            
          }
        }






        }

        //// Modal Command with pickfirst selection
        //[CommandMethod("MyGroup", "MyPickFirst", "MyPickFirstLocal", CommandFlags.Modal | CommandFlags.UsePickSet)]
        //public void MyPickFirst() // This method can have any name
        //{
        //    PromptSelectionResult result = Application.DocumentManager.MdiActiveDocument.Editor.GetSelection();
        //    if (result.Status == PromptStatus.OK)
        //    {
        //        // There are selected entities
        //        // Put your command using pickfirst set code here
        //    }
        //    else
        //    {
        //        // There are no selected entities
        //        // Put your command code here
        //    }
        //}

        //// Application Session Command with localized name
        //[CommandMethod("MyGroup", "MySessionCmd", "MySessionCmdLocal", CommandFlags.Modal | CommandFlags.Session)]
        //public void MySessionCmd() // This method can have any name
        //{
        //    // Put your command code here
        //}

        //// LispFunction is similar to CommandMethod but it creates a lisp 
        //// callable function. Many return types are supported not just string
        //// or integer.
        //[LispFunction("MyLispFunction", "MyLispFunctionLocal")]
        //public int MyLispFunction(ResultBuffer args) // This method can have any name
        //{
        //    // Put your command code here

        //    // Return a value to the AutoCAD Lisp Interpreter
        //    return 1;
        //}

    }

}
