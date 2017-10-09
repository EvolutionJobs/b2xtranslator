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
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.BiffView
{
    public partial class MainWindow : Form
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void cmdBrowse_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtFileName.Text = openFileDialog.FileName;
                cmdCreate.Focus();
            }
        }

        private void txtFileName_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists(txtFileName.Text))
                {
                    cmdCreate.Enabled = true;
                    toolStripStatusLabel1.Text = "Ready";
                }
                else
                {
                    cmdCreate.Enabled = false;
                    toolStripStatusLabel1.Text = "Please select a file.";
                }
                toolStripProgressBar1.Value = 0;
            }
            catch
            {
                cmdCreate.Enabled = false;
            }
        }

        private void cmdExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmdCreate_Click(object sender, EventArgs e)
        {
            if (!this.backgroundWorker.IsBusy)
            {
                BiffViewerOptions options = new BiffViewerOptions();
                options.ShowErrors = true;
                options.UseTempFolder = true;
                options.OutputFileName = options.GetTempFileName(txtFileName.Text);
                options.ShowInBrowser = true;
                options.InputDocument = txtFileName.Text;

                this.toolStripProgressBar1.Value = 0;

                backgroundWorker.RunWorkerAsync(options);

                // Disable the button for the duration of the download.
                this.cmdCreate.Text = "Cancel";
            }
            else
            {
                this.cmdCreate.Enabled = false;
                this.backgroundWorker.CancelAsync();
            }
        }

        private void linkWeb_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkWeb.LinkVisited = true;
            Util.VisitLink(Util.ProjectWebSite);
        }

        private void linkAbout_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new About().Show();
        }

        void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            // Get the BackgroundWorker that raised this event.
            BackgroundWorker worker = sender as BackgroundWorker;

            BiffViewerOptions options = e.Argument as BiffViewerOptions;
            BiffViewer viewer = new BiffViewer(options);
            try
            {
                viewer.DoTheMagic(worker);
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message, "BiffView++", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception)
            {
                throw;
            }
            
            if (worker.CancellationPending)
            {
                e.Cancel = true;
            }
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.toolStripProgressBar1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // First, handle the case where an exception was thrown.
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message, "BiffView++", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (e.Cancelled)
            {
                this.toolStripStatusLabel1.Text = "Cancelled";
            }
            else
            {
                this.toolStripStatusLabel1.Text = "Completed";
            }

            this.cmdCreate.Enabled = true;
            this.cmdCreate.Text = "Create";
        }

        private void toolStripStatusLabelAbout_Click(object sender, EventArgs e)
        {
            new About().Show();
        }

        private void MainWindow_Shown(object sender, EventArgs e)
        {
            cmdBrowse.Focus();
        }
    }
}