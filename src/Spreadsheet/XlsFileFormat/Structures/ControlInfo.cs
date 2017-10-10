/*
 * Copyright (c) 2009, DIaLOGIKa
 *
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without 
 * modification, are permitted provided that the following conditions are met:
 * 
 *     * Redistributions of source code must retain the above copyright 
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright 
 *       notice, this list of conditions and the following disclaimer in the 
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the names of copyright holders, nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software 
 *       without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
 * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
 * OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
 */

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies the properties of some form control in a Dialog Sheet. 
    /// 
    /// The control MUST be a group, radio button, label, button or checkbox.
    /// </summary>
    public class ControlInfo
    {
        /// <summary>
        /// A bit that specifies whether this control dismisses the Dialog Sheet and performs the default behavior. 
        /// 
        /// If the control is not a button, the value MUST be 0.
        /// </summary>
        public bool fDefault;

        /// <summary>
        /// A bit that specifies whether this control is intended to load context-sensitive help for the Dialog Sheet. 
        /// 
        /// If the control is not a button, the value MUST be 0.
        /// </summary>
        public bool fHelp;

        /// <summary>
        /// A bit that specifies whether this control dismisses the Dialog Sheet and take no action. 
        /// 
        /// If the control is not a button, the value MUST be 0.
        /// </summary>
        public bool fCancel;

        /// <summary>
        /// A bit that specifies whether this control dismisses the Dialog Sheet. 
        /// 
        /// If the control is not a button, the value MUST be 0.
        /// </summary>
        public bool fDismiss;

        /// <summary>
        /// A signed integer that specifies the Unicode character of the control‘s accelerator key. 
        /// 
        /// The value MUST be greater than or equal to 0x0000. A value of 0x0000 specifies there is no accelerator associated with this control.
        /// </summary>
        public short accel1;

        public ControlInfo(IStreamReader reader)
        {
            var flags = reader.ReadUInt16();
            this.fDefault = Utils.BitmaskToBool(flags, 0x0001);
            this.fHelp = Utils.BitmaskToBool(flags, 0x0002);
            this.fCancel = Utils.BitmaskToBool(flags, 0x0004);
            this.fDismiss = Utils.BitmaskToBool(flags, 0x0008);

            this.accel1 = reader.ReadInt16();

            reader.ReadBytes(2);
        }
    }
}
