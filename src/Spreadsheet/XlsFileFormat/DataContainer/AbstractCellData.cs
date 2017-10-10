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

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// Abstract class which stores some data
    /// </summary>
    public abstract class AbstractCellData: IComparable
    {
        /// Attributes ///
        
        /// <summary>
        /// Row of the Object
        /// </summary>
        private int row;
        /// <summary>
        /// Getter Setter from Row 
        /// </summary>
	    public int Row
	    {
		    get { return row;}
		    set { row = value;}
	    }
	
        /// <summary>
        /// The column of the object 
        /// </summary>
        private int col;
        /// <summary>
        /// Getter Setter from col 
        /// </summary>
	    public int Col
	    {
		    get { return col;}
		    set { col = value;}
	    }

        /// <summary>
        /// TemplateID from this object 
        /// References to a template field 
        /// </summary>
        private int templateID;
        /// <summary>
        /// Getter setter from the templateID attribute 
        /// </summary>
        public int TemplateID
        {
            get { return templateID; }
            set { templateID = value; }
        }


        /// Constructors ///

        /// <summary>
        /// Ctor 
        /// </summary>
        public AbstractCellData() : this (0,0,0)  { }

        /// <summary>
        /// ctor
        /// </summary>
        /// <param name="row">Rownumber of the object</param>
        /// <param name="col">Colnumber of the object</param>
        /// <param name="templateID">ID of the objectstyletemplate </param>
        public AbstractCellData(int row, int col, int templateID)
        {
            this.row = row;
            this.col = col;
            this.templateID = templateID; 
        }

        /// Abstract Methods ///

        /// <summary>
        /// Returns a String from the stored Value
        /// </summary>
        /// <returns></returns>
        public abstract string getValue();

        /// <summary>
        /// Sets the value 
        /// </summary>
        /// <param name="obj"></param>
        public abstract void setValue(Object obj);

        /// <summary>
        /// Implements the compareble interface 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        int IComparable.CompareTo(object obj)
        {
            var cell = (AbstractCellData)obj;
            if (this.col > cell.col)
                return (1);
            if (this.col < cell.col)
                return (-1);
            else
                return (0);
        }
        
    }
}
