using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using Microsoft.Win32;

namespace DIaLOGIKa.b2xtranslator.xls2x
{
    [RunInstaller(true)]
    public partial class CustomInstaller : Installer
    {
        private string creationMark = "CreatedBy";
        private string creationValue = "xls2xInstaller";

        public CustomInstaller()
        {
            InitializeComponent();
        }

        public override void Install(IDictionary savedState)
        {
            base.Install(savedState);

            if (Context.Parameters["createcontextmenuxls"] == "1")
            {
                addContextMenuEntry();
            }
        }

        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);

            removeContextMenuEntry();
        }

        /// <summary>
        /// Creates the context menu entry
        /// </summary>
        private void addContextMenuEntry()
        {
            RegistryKey shellKey = getShellKey(Program.ContextMenuInputExtension);
            if (shellKey != null)
            {
                // create the context menu entry
                RegistryKey entryKey = shellKey.CreateSubKey(Program.ContextMenuText);

                // set the installer's mark
                entryKey.SetValue(creationMark, creationValue);

                // create command subkey
                RegistryKey convertCommand = entryKey.CreateSubKey("Command");

                // set the ppt path as value
                convertCommand.SetValue("", String.Format("\"{0}\" \"%1\"", Context.Parameters["assemblypath"]));
            }
        }

        /// <summary>
        /// Removes the context menu entry if it is present 
        /// and if it was created by the installer.
        /// </summary>
        private void removeContextMenuEntry()
        {
            try
            {
                RegistryKey shellKey = getShellKey(Program.ContextMenuInputExtension);
                if (shellKey != null)
                {
                    // if the entry is present and was created by the installer
                    RegistryKey entryKey = shellKey.OpenSubKey(Program.ContextMenuText);
                    if (entryKey != null)
                    {
                        object mark = entryKey.GetValue(creationMark);
                        if (mark != null && (string)mark == creationValue)
                        {
                            shellKey.DeleteSubKeyTree(Program.ContextMenuText);
                        }
                    }
                }
            }
            catch (Exception) { }
        }

        /// <summary>
        /// Returns the shell key that is related to the given extension
        /// </summary>
        /// <param name="triggerExtension"></param>
        /// <returns></returns>
        private RegistryKey getShellKey(string triggerExtension)
        {
            RegistryKey result = null;
            try
            {
                string defaultApp = (string)Registry.ClassesRoot.OpenSubKey(triggerExtension).GetValue("");
                result = Registry.ClassesRoot.CreateSubKey(defaultApp).CreateSubKey("shell");
            }
            catch (Exception)
            {
            }
            return result;
        }
    }
}
