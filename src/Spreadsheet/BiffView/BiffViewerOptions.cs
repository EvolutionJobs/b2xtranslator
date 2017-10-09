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
using System.IO;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.BiffView
{
    public enum BiffViewerMode
    {
        Console,
        File
    }

    public class BiffViewerOptions
    {
        private string _tmpFileName = "";
        private string _tmpWorkFolder = "";

        ~BiffViewerOptions()
        {
            // cleanup temp folder
            if (!String.IsNullOrEmpty(_tmpFileName))
            {
                File.Delete(_tmpFileName);
            }

            if (!String.IsNullOrEmpty(_tmpWorkFolder))
            {
                Directory.Delete(_tmpWorkFolder, true);
            }
        }

        private long _startPosition = -1;
        public long StartPosition
        {
            get { return _startPosition; }
            set { _startPosition = value; }
        }

        private long _endPosition = -1;
        public long EndPosition
        {
            get { return _endPosition; }
            set { _endPosition = value; }
        }

        private bool _showErrors = false;
        public bool ShowErrors
        {
            get { return _showErrors; }
            set { _showErrors = value; }
        }

        private bool _useOnlineSpecification = false;
        public bool UseOnlineSpecification
        {
            get { return _useOnlineSpecification; }
            set { _useOnlineSpecification = value; }
        }

        private bool _useTempFolder = false;
        public bool UseTempFolder
        {
            get { return _useTempFolder; }
            set { _useTempFolder = value; }
        }

        private bool _printTextOnly = false;
        public bool PrintTextOnly
        {
            get { return _printTextOnly; }
            set { _printTextOnly = value; }
        }

        private bool _showInBrowser = false;
        public bool ShowInBrowser
        {
            get { return _showInBrowser; }
            set { _showInBrowser = value; }
        }

        private string _outputFileName;
        public string OutputFileName
        {
            get { return _outputFileName; }
            set { _outputFileName = value; }
        }

        private string _inputDocument;
        public string InputDocument
        {
            get { return _inputDocument; }
            set { _inputDocument = value; }
        }


        public BiffViewerMode Mode
        {
            get
            {
                if (!String.IsNullOrEmpty(this.OutputFileName))
                {
                    return BiffViewerMode.File;
                }
                return BiffViewerMode.Console;
            }
        }

        public string GetTempFileName(string inputDocument)
        {
            if (String.IsNullOrEmpty(_tmpFileName))
            {
                // WARNING!!!
                // only use a temporary folder
                // _tmpWorkFolder and all its contents is deleted in the destructor!!!
                do
                {
                    _tmpWorkFolder = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                }
                while (Directory.Exists(_tmpWorkFolder));
                Directory.CreateDirectory(_tmpWorkFolder);

                FileInfo fi = new FileInfo(inputDocument);
                if (this.PrintTextOnly)
                {
                    _tmpFileName = Path.Combine(_tmpWorkFolder, fi.Name + ".txt");
                }
                else
                {
                    _tmpFileName = Path.Combine(_tmpWorkFolder, fi.Name + ".html");
                }
            }
            return _tmpFileName;
        }
    }
}
