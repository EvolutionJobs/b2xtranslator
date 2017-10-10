using System;
using System.IO;

namespace b2xtranslator.ZipUtils
{
    class ZlibZipReader : ZipReader, IDisposable
    {
        private IntPtr handle;
        private bool disposed = false;

        internal ZlibZipReader(string path)
        {
            string resolvedPath = ZipLib.ResolvePath(path);
            if (!File.Exists(resolvedPath))
            {
                throw new FileNotFoundException("File does not exist:" + path);
            }
            this.handle = ZipLib.unzOpen(resolvedPath);
            if (this.handle == IntPtr.Zero)
            {
                throw new ZipException("Unable to open ZIP file:" + path);
            }
        }

        ~ZlibZipReader() {
            Dispose(false);
        }

        public override void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing) {
            if (!this.disposed) {
                if (disposing) {
                    // Dispose managed resources : none here
                }
                Close();
            }
        }

        public override void Close()
        {
            if (this.handle != IntPtr.Zero) {
                int result = ZipLib.unzClose(this.handle);
                this.handle = IntPtr.Zero;
                // Question: should we raise this exception ?
                if (result != 0) {
                    throw new ZipException("Error closing file - Errorcode: " + result);
                }
            }
        }

        public override Stream GetEntry(string relativePath)
        {
            string resolvedPath = ZipLib.ResolvePath(relativePath);
            if (ZipLib.unzLocateFile(this.handle, resolvedPath, 0) != 0)
            {
                throw new ZipEntryNotFoundException("Entry not found:" + relativePath);
            }

            var entryInfo = new ZipEntryInfo();
            int result = ZipLib.unzGetCurrentFileInfo(this.handle, out entryInfo, null, UIntPtr.Zero, null, UIntPtr.Zero, null, UIntPtr.Zero);
            if (result != 0)
            {
                throw new ZipException("Error while reading entry info: " + relativePath + " - Errorcode: " + result);
            }

            result = ZipLib.unzOpenCurrentFile(this.handle);
            if (result != 0)
            {
                throw new ZipException("Error while opening entry: " + relativePath + " - Errorcode: " + result);
            }

            var buffer = new byte[entryInfo.UncompressedSize.ToUInt64 ()];
            int bytesRead = 0;
            if ((bytesRead = ZipLib.unzReadCurrentFile(this.handle, buffer, (uint)entryInfo.UncompressedSize)) < 0)
            {
                throw new ZipException("Error while reading entry: " + relativePath + " - Errorcode: " + result);
            }

            result = ZipLib.unzCloseCurrentFile(this.handle);
            if (result != 0)
            {
                throw new ZipException("Error while closing entry: " + relativePath + " - Errorcode: " + result);
            }

            return new MemoryStream(buffer, 0, bytesRead);
        }
    }
}
