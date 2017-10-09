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
using System.Windows.Forms;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.BiffView
{
    public partial class About : Form
    {
        public About()
        {
            InitializeComponent();

            this.Text = String.Format("About {0}", Util.AssemblyTitle);
            this.labelProductName.Text = Util.AssemblyProduct;
            this.labelVersion.Text = String.Format("Version {0}", Util.AssemblyVersion);
            this.linkWeb.Text = String.Format("{0} {1}", Util.AssemblyCopyright, Util.ProjectWebSite);
            this.linkWeb.LinkArea = new System.Windows.Forms.LinkArea(Util.AssemblyCopyright.Length + 1, this.linkWeb.Text.Length);
            this.linkCompany.Text = String.Format("Developed by {0}", Util.AssemblyCompany);
            this.linkCompany.LinkArea = new System.Windows.Forms.LinkArea("Developed at ".Length, Util.AssemblyCompany.Length);
        }

        private void linkCompany_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkCompany.LinkVisited = true;
            Util.VisitLink(Util.DiaWebSite);
        }

        private void linkWeb_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkWeb.LinkVisited = true;
            Util.VisitLink(Util.ProjectWebSite);
        }

        
        private void cmdOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        
    }
}