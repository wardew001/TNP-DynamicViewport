using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using System;

[assembly: CommandClass(typeof(TNPPlugin.Plugin))]

namespace TNPPlugin
{
    public class Plugin: IExtensionApplication
    {
        public void Initialize(){}

        public void Terminate() {}
    }
}
