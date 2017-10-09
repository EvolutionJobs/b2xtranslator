/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using Microsoft.Win32;
using DIaLOGIKa.b2xtranslator.Shell;
using System.Reflection;
using System.Windows.Forms;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.doc2x
{
    [RunInstaller(true)]
    public partial class CustomInstaller : Installer
    {
        private string creationMark = "CreatedBy";
        private string creationValue = "doc2xInstaller";


        public CustomInstaller()
        {
            InitializeComponent();
        }


        public override void Install(IDictionary savedState)
        {
            base.Install(savedState);

            if (Context.Parameters["createcontextmenudoc"] == "1")
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
                            // remove the entry
                            shellKey.DeleteSubKeyTree(Program.ContextMenuText);
                        }
                    }
                }
            }
            catch (Exception) { 
            }
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
