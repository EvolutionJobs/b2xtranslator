using System;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.ZipUtils
{
    class ZlibZipWriter : ZipWriter
    {
        private IntPtr handle;
        private bool entryOpened;

        internal ZlibZipWriter(string path)
        {
            this.Filename = path;
            string resolvedPath = ZipLib.ResolvePath(path);
            this.handle = ZipLib.zipOpen(resolvedPath, 0);
            if (this.handle == IntPtr.Zero)
            {
                // Small trick to get exact error message...
                try {
                    using (var writer = File.Create(path)) {
                        writer.WriteByte(0);
                    }
                    File.Delete(path);
                    throw new ZipCreationException();
                } catch (Exception ex) {
                    throw new ZipCreationException(ex.Message);
                }
            }
            this.entryOpened = false;
		}
	    
	    public override void AddEntry(string relativePath)
        {
            string resolvedPath = ZipLib.ResolvePath(relativePath);
            ZipFileEntryInfo info;
            info.DateTime = DateTime.Now;

            if (this.entryOpened)
            {
                ZipLib.zipCloseFileInZip(this.handle);
                this.entryOpened = false;
            }

            int result = ZipLib.zipOpenNewFileInZip(this.handle, resolvedPath, out info, null, 0, null, 0, string.Empty, (int)CompressionMethod.Deflated, (int)CompressionLevel.Default);
            if (result < 0)
            {
                throw new ZipException("Error while opening entry for writing: " + relativePath + " - Errorcode: " + result);
            }
            this.entryOpened = true;
	    }	    	
	    
	    public override void Close()
        {
            if (this.handle != IntPtr.Zero) {
                int result;
                if (this.entryOpened) {
                    result = ZipLib.zipCloseFileInZip(this.handle);
                    if (result != 0) {
                        throw new ZipException("Error while closing entry - Errorcode: " + result);
                    }
                    this.entryOpened = false;
                }
                result = ZipLib.zipClose(this.handle, "");
                this.handle = IntPtr.Zero;
                // Should we raise this exception ?
                if (result != 0) {
                    throw new ZipException("Error while closing ZIP file - Errorcode: " + result);
                }
            }
        }
	    
	    public override int Read(byte[] buffer, int offset, int count)
		{
            throw new ZipException("Error, Read not allowed on this stream");
		}

		public override void Write(byte[] buffer, int offset, int count)
		{
            int result;
            if (offset != 0)
            {
                var newBuffer = new byte[count];
                Array.Copy(buffer, offset, newBuffer, 0, count);
                result = ZipLib.zipWriteInFileInZip(this.handle, newBuffer, (uint)count);
            }
            else
            {
                result = ZipLib.zipWriteInFileInZip(this.handle, buffer, (uint)count);
            }

            if (result < 0)
            {
                throw new ZipException("Error while writing entry - Errorcode: " + result);
            }
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return 0;
        }

        public override void SetLength(long value)
        {
            return;
        }

        public override void Flush()
        {
            return;
        }
		
		public override long Position {
			get {
                throw new ZipException("Position not available on this stream");
			}
			set {
                throw new ZipException("Position not available on this stream");
			}
		}
		
		public override long Length {
			get {
				return 0;
			}
		}
		
		public override bool CanRead {
			get {
				return false;
			}
		}
		
		public override bool CanWrite {
			get {
				return true;
			}
		}
		
		public override bool CanSeek {
			get {
				return false;
			}
		}
    }
}
