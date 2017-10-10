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

using System.Collections.Generic;
using DIaLOGIKa.b2xtranslator.Tools; 

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer
{
    /// <summary>
    /// This class stores the rowdata from a specific row 
    /// </summary>
    public class RowData
    {
        /// <summary>
        /// The row number 
        /// </summary>
        private int row;
        public int Row
        {
            get { return this.row; }
            set { this.row = value; }
        }
        
        /// <summary>
        /// Collection of cellobjects 
        /// </summary>
        private List<AbstractCellData> cells;
        public List<AbstractCellData> Cells
        {
            get { return this.cells; }
            set { this.cells = value; }
        }

        public TwipsValue height;
        public bool hidden;
        public int outlineLevel;
        public bool collapsed;
        public bool customFormat;
        public int style;
        public bool thickBot;
        public bool thickTop;
        public bool customHeight;

        public int minSpan;
        public int maxSpan; 
        /// <summary>
        /// Ctor 
        /// </summary>
        public RowData()
            : this(0)
        {
            this.outlineLevel = -1;
            this.minSpan = -1;
            this.maxSpan = -1; 
        }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="row">Rowid</param>
        public RowData(int row)
        {
            this.row = row;
            this.cells = new List<AbstractCellData>(); 
        }

        /// <summary>
        /// Add a cellobject to the collection 
        /// </summary>
        /// <param name="cell">Cellobject</param>
        public void addCell(AbstractCellData cell)
        {
            if (!this.checkCellExists(cell))
                this.cells.Add(cell); 
        }

        /// <summary>
        /// method checks if a cell exists or not 
        /// </summary>
        /// <param name="cell"></param>
        /// <returns>true if the cell exists and false if not</returns>
        public bool checkCellExists(AbstractCellData cell)
        {
            foreach (var var in this.cells)
            {
                if (var.Col == cell.Col)
                {
                    return true; 
                }
            }
            return false; 
        }
    }
}
