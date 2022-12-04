// (C) Copyright 2022 by  
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Interop;
using System;

// This line is not mandatory, but improves loading performances
[assembly: ExtensionApplication(typeof(AutoCAD_CSharp_plug_in2.MyPlugin))]

namespace AutoCAD_CSharp_plug_in2
{
    // This class is instantiated by AutoCAD once and kept alive for the 
    // duration of the session. If you don't do any one time initialization 
    // then you should remove this class.
    public class MyPlugin : IExtensionApplication
    {

        void IExtensionApplication.Initialize()
        {

        }

        void IExtensionApplication.Terminate()
        {

        }

    }

}
