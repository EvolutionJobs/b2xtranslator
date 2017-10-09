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
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public enum RecordNumber : ushort
    {
        DATE1904 = 0x22, 	    // 1904 Date System
        ADDIN = 0x87, 	        // Workbook Is an Add-in Macro
        ADDMENU = 0xC2, 	    // Menu Addition
        ARRAY = 0x221, 	        // Array-Entered Formula
        AUTOFILTER = 0x9E, 	    // AutoFilter Data
        AUTOFILTER12 = 0x87E, 	// AutoFilter Data Introduced in Excel 2007
        AUTOFILTERINFO = 0x9D, 	// Drop-Down Arrow Count
        AUTOWEBPUB = 0x8C0,     // Auto web publish storage
        BACKUP = 0x40, 	        // Save Backup Version of the File
        BLANK = 0x201, 	        // Cell Value, Blank Cell
        BOF = 0x809, 	        // Beginning of File
        BOOKBOOL = 0xDA, 	    // Workbook Option Flag
        BOOKEXT = 0x863, 	    //  Extra Book Info
        BOOLERR = 0x205, 	    // Cell Value, Boolean or Error
        BOTTOMMARGIN = 0x29, 	// Bottom Margin Measurement
        BOUNDSHEET = 0x85, 	    // Sheet Information
        CALCCOUNT = 0x0C, 	    // Iteration Count
        CALCMODE = 0x0D, 	    // Calculation Mode
        CELLWATCH = 0x86C, 	    //  Cell Watch
        CF = 0x1B1, 	        // Conditional Formatting Conditions
        CF12 = 0x87A, 	        // Conditional Formatting Condition 12
        CFEX = 0x87B, 	        // Conditional Formatting Extension
        CODENAME = 0x42, 	    // VBE Object Name
        CODEPAGE = 0x42, 	    // Default Code Page
        COLINFO = 0x7D, 	    // Column Formatting Information
        COMPAT12 = 0x88C, 	    // Compatibility Checker 12
        COMPRESSPICTURES = 0x89B, 	// Automatic Picture Compression Mode
        CONDFMT = 0x1B0, 	    // Conditional Formatting Range Information
        CONDFMT12 = 0x879, 	    // Conditional Formatting Range Information 12
        CONTINUE = 0x3C, 	    // Continues Long Records
        CONTINUEFRT = 0x812, 	// Continued FRT
        CONTINUEFRT11 = 0x875, 	//  Continue FRT 11
        CONTINUEFRT12 = 0x87F, 	// Continued FRT 12
        COORDLIST = 0xA9, 	    // Polygon Object Vertex Coordinates
        COUNTRY = 0x8C, 	    // Default Country and WIN.INI Country
        CRASHRECERR = 0x865, 	//  Crash Recovery Error
        CRN = 0x5A, 	        // Nonresident Operands
        CRTCOOPT = 0x8cb, 	    //  Color options for Chart series in Mac Office 11
        DATALABEXT = 0x86A, 	//  Chart Data Label Extension
        DATALABEXTCONTENTS = 0x86B, 	//  Chart Data Label Extension Contents
        DBCELL = 0xD7, 	        // Stream Offsets
        DBQUERYEXT = 0x803, 	// Database Query Extensions
        DCON = 0x50, 	        // Data Consolidation Information
        DCONBIN = 0x1B5, 	    // Data Consolidation Information
        DCONN = 0x876, 	        // Data Connection
        DCONNAME = 0x52, 	    // Data Consolidation Named References
        DCONREF = 0x51, 	    // Data Consolidation References
        DEFAULTROWHEIGHT = 0x225, 	// Default Row Height
        DEFCOLWIDTH = 0x55, 	// Default Width for Columns
        DELMENU = 0xC3, 	    // Menu Deletion
        DELTA = 0x10, 	        // Iteration Increment
        DIMENSIONS = 0x200, 	// Cell Table Size
        DOCROUTE = 0xB8, 	    // Routing Slip Information
        DROPDOWNOBJIDS = 0x874, //  Drop Down Object
        DSF = 0x161, 	        // Double Stream File
        DV = 0x1BE, 	        // Data Validation Criteria
        DVAL = 0x1B2,       	// Data Validation Information
        DXF = 0x88D, 	        // Differential XF
        EDG = 0x88, 	        // Edition Globals
        EOF = 0x0A, 	        // End of File
        EXCEL9FILE = 0x1C0, 	// Excel 9 File
        EXTERNCOUNT = 0x16, 	// Number of External References
        EXTERNNAME = 0x23, 	    // Externally Referenced Name
        EXTERNSHEET = 0x17, 	// External Reference
        EXTSST = 0xFF, 	        // Extended Shared String Table
        EXTSTRING = 0x804, 	    // FRT String
        FEAT = 0x868, 	        //  Shared Feature Record
        FEAT11 = 0x872, 	    //  Shared Feature 11 Record
        FEAT12 = 0x878, 	    //  Shared Feature 12 Record
        FEATHEADR = 0x867, 	    //  Shared Feature Header
        FEATHEADR11 = 0x871, 	//  Shared Feature Header 11
        FEATINFO = 0x86d, 	    //  Shared Feature Info Record
        FEATINFO11 = 0x873, 	//  Shared Feature Info 11 Record
        FILEPASS = 0x2F, 	    // File Is Password-Protected
        FILESHARING = 0x5B, 	// File-Sharing Information
        FILESHARING2 = 0x1A5, 	// File-Sharing Information for Shared Lists
        FILTERMODE = 0x9B, 	    // Sheet Contains Filtered List
        FMQRY = 0x8c6, 	        //  Filemaker queries
        FMSQRY = 0x8c7, 	    //  File maker queries
        FNGRP12 = 0x898, 	    // Function Group
        FNGROUPCOUNT = 0x9C, 	// Built-in Function Group Count
        FNGROUPNAME = 0x9A, 	// Function Group Name
        FONT = 0x231,        	// Font Description

        // This is a deviation from the specification
        FONT2 = 0x31, 	        // Font Description

        FOOTER = 0x15, 	        // Print Footer on Each Page
        FORCEFULLCALCULATION = 0x8A3, 	// Force Full Calculation Mode
        FORMAT = 0x41E, 	    // Number Format
        // FORMULA = 0x406, 	// Cell Formula
        FORMULA = 0x06, 	    // Cell Formula
        GCW = 0xAB, 	        // Global Column-Width Flags
        GRIDSET = 0x82, 	    // State Change of Gridlines Option
        GUIDTYPELIB = 0x897, 	// VB Project Typelib GUID
        GUTS = 0x80, 	        // Size of Row and Column Gutters
        HCENTER = 0x83, 	    // Center Between Horizontal Margins
        HEADER = 0x14, 	        // Print Header on Each Page
        HEADERFOOTER = 0x89C, 	//  Header Footer
        HFPicture = 0x866, 	    //  Header / Footer Picture
        HIDEOBJ = 0x8D, 	    // Object Display Options
        HLINK = 0x1B8, 	        // Hyperlink
        HLINKTOOLTIP = 0x800, 	// Hyperlink Tooltip
        HORIZONTALPAGEBREAKS = 0x1B, 	// Explicit Row Page Breaks
        IMDATA = 0x7F, 	        // Image Data
        INDEX = 0x20B, 	        // Index Record
        INTERFACEEND = 0xE2, 	// End of User Interface Records
        INTERFACEHDR = 0xE1, 	// Beginning of User Interface Records
        ITERATION = 0x11, 	    // Iteration Mode
        LABEL = 0x204, 	        // Cell Value, String Constant
        LABELSST = 0xFD, 	    // Cell Value, String Constant/SST
        LEFTMARGIN = 0x26, 	    // Left Margin Measurement
        LHNGRAPH = 0x95, 	    // Named Graph Information
        LHRECORD = 0x94, 	    // .WK? File Conversion Information
        LIST12 = 0x877, 	    //  Extra Table Data Introduced in Excel 2007
        LISTCF = 0x8c5, 	    //  List Cell Formatting
        LISTCONDFMT = 0x8c4, 	//  List Conditional Formatting
        LISTDV = 0x8c3, 	    //  List Data Validation
        LISTFIELD = 0x8c2,  	//  List Field
        LISTOBJ = 0x8c1, 	    //  List Object
        LNEXT = 0x8c9,      	//  Extension information for borders in Mac Office 11
        LPR = 0x98, 	        // Sheet Was Printed Using LINE.PRINT()
        MDTB = 0x88A, 	        // Block of Metadata Records
        MDTINFO = 0x884, 	    // Information about a Metadata Type
        MDXPROP = 0x888,    	// Member Property MDX Metadata
        MDXKPI = 0x889, 	    // Key Performance Indicator MDX Metadata
        MDXSET = 0x887, 	    // Set MDX Metadata
        MDXSTR = 0x885, 	    // MDX Metadata String
        MDXTUPLE = 0x886, 	    // Tuple MDX Metadata
        MERGECELLS = 0xE5,  	// Merged Cells
        MKREXT = 0x8ca, 	    //  Extension information for markers in Mac Office 11
        MMS = 0xC1, 	        // ADDMENU/DELMENU Record Group Count
        MSODRAWING = 0xEC,  	// Microsoft Office Drawing
        MSODRAWINGGROUP = 0xEB, // Microsoft Office Drawing Group
        MSODRAWINGSELECTION = 0xED, 	// Microsoft Office Drawing Selection
        MTRSETTINGS = 0x89A, 	// Multi-Threaded Calculation Settings
        MULBLANK = 0xBE, 	    // Multiple Blank Cells
        MULRK = 0xBD, 	        // Multiple RK Cells
        NAME = 0x18, 	        // Defined Name
        NAMECMT = 0x894, 	    // Name Comment
        NAMEFNGRP12 = 0x899, 	// Extra Function Group
        NAMEPUBLISH = 0x893, 	// Publish to Excel Server Data for Name
        NOTE = 0x1C, 	        // Comment Associated with a Cell
        NUMBER = 0x203,     	// Cell Value, Floating-Point Number
        OBJ = 0x5D, 	        // Describes a Graphic Object
        OBJPROTECT = 0x63,  	// Objects Are Protected
        OBPROJ = 0xD3, 	        // Visual Basic Project
        OLEDBCONN = 0x80A,  	// OLE Database Connection
        OLESIZE = 0xDE, 	    // Size of OLE Object
        PALETTE = 0x92, 	    // Color Palette Definition
        PANE = 0x41, 	        // Number of Panes and Their Position
        PARAMQRY = 0xDC, 	    // Query Parameters
        PASSWORD = 0x13, 	    // Protection Password
        PLS = 0x4D, 	        // Environment-Specific Print Record
        PLV = 0x8c8, 	        //  Page Layout View in Mac Excel 11
        PLV12 = 0x88B, 	        //  Page Layout View Settings in Excel 2007
        PRECISION = 0x0E, 	    // Precision
        PRINTGRIDLINES = 0x2B, 	// Print Gridlines Flag
        PRINTHEADERS = 0x2A, 	// Print Row/Column Labels
        PROTECT = 0x12, 	    // Protection Flag
        PROT4REV = 0x1AF, 	    // Shared Workbook Protection Flag
        PROT4REVPASS = 0x1BC, 	//  Shared Workbook Protection Password
        PUB = 0x89, 	        //  Publisher
        QSI = 0x1AD, 	        // External Data Range
        QSIF = 0x807, 	        // Query Table Field Formatting
        QSIR = 0x806, 	        // Query Table Formatting
        QSISXTAG = 0x802, 	    // PivotTable and Query Table Extensions
        REALTIMEDATA = 0x813, 	//  Real-Time Data (RTD)
        RECALCID = 0x1C1, 	    // Recalc Information
        RECIPNAME = 0xB9, 	    // Recipient Name
        REFMODE = 0x0F, 	    // Reference Mode
        REFRESHALL = 0x1B7, 	// Refresh Flag
        RIGHTMARGIN = 0x27, 	// Right Margin Measurement
        RK = 0x27E, 	        // Cell Value, RK Number
        ROW = 0x208, 	        // Describes a Row
        RSTRING = 0xD6, 	    // Cell with Character Formatting
        SAVERECALC = 0x5F, 	    // Recalculate Before Save
        SCENARIO = 0xAF, 	    // Scenario Data
        SCENMAN = 0xAE, 	    // Scenario Output Data
        SCENPROTECT = 0xDD, 	// Scenario Protection
        SCL = 0xA0,          	// Window Zoom Magnification
        SELECTION = 0x1D, 	    // Current Selection
        SETUP = 0xA1, 	        // Page Setup
        SHAPEPROPSSTREAM = 0x08A4,  // Shape formatting properties for chart elements
        SHEETEXT = 0x862,    	//  Extra Sheet Info
        SHRFMLA = 0x4BC,     	// Shared Formula
        SORT = 0x90, 	        // Sorting Options
        SORTDATA12 = 0x895, 	// Sort Data 12
        SOUND = 0x96, 	        // Sound Note
        SST = 0xFC, 	        // Shared String Table
        STANDARDWIDTH = 0x99, 	// Standard Column Width
        STRING = 0x207, 	    // String Value of a Formula
        STYLE = 0x293, 	        // Style Information
        STYLEEXT = 0x892, 	    // Named Cell Style Extension
        SUB = 0x91, 	        // Subscriber
        SUPBOOK = 0x1AE, 	    // Supporting Workbook
        SXADDL = 0x864, 	    //  Pivot Table Additional Info
        SXADDL12 = 0x881, 	    // Additional Workbook Connections Information
        SXDB = 0xC6, 	        // PivotTable Cache Data
        SXDBEX = 0x122, 	    // PivotTable Cache Data
        SXDI = 0xC5, 	        // Data Item
        SXDXF = 0xF4, 	        //  Pivot Table Formatting
        SXEX = 0xF1, 	        // PivotTable View Extended Information
        SXEXT = 0xDC, 	        // External Source Information
        SXFDBTYPE = 0x1BB, 	    // SQL Datatype Identifier
        SXFILT = 0xF2, 	        // PivotTable Rule Filter
        SXFMLA = 0xF9, 	        //  Pivot Table Parsed Expression
        SXFORMAT = 0xFB, 	    // PivotTable Format Record
        SXFORMULA = 0x103, 	    // PivotTable Formula Record
        SXIDSTM = 0xD5, 	    // Stream ID
        SXITM = 0xF5, 	        //  Pivot Table Item Indexes
        SXIVD = 0xB4, 	        // Row/Column Field IDs
        SXLI = 0xB5, 	        // Line Item Array
        SXNAME = 0xF6, 	        // PivotTable Name
        SXPAIR = 0xF8, 	        // PivotTable Name Pair
        SXPI = 0xB6, 	        // Page Item
        SXPIEX = 0x80E, 	    // OLAP Page Item Extensions
        SXRULE = 0xF0, 	        // PivotTable Rule Data
        SXSELECT = 0xF7, 	    // PivotTable Selection Information
        SXSTRING = 0xCD, 	    // String
        SXTBL = 0xD0, 	        // Multiple Consolidation Source Info
        SXTBPG = 0xD2, 	        // Page Item Indexes
        SXTBRGIITM = 0xD1, 	    // Page Item Name Count
        SXTH = 0x80D, 	        // PivotTable OLAP Hierarchy
        SXVD = 0xB1, 	        // View Fields
        SXVDEX = 0x100, 	    // Extended PivotTable View Fields
        SXVDTEX = 0x80F, 	    // View Dimension OLAP Extensions
        SXVI = 0xB2, 	        // View Item
        SXVIEW = 0xB0, 	        // View Definition
        SXVIEWEX = 0x80C, 	    // Pivot Table OLAP Extensions
        SXVIEWEX9 = 0x810, 	    // Pivot Table Extensions
        SXVS = 0xE3, 	        // View Source
        TABID = 0x13D, 	        // Sheet Tab Index Array
        TABIDCONF = 0xEA, 	    // Sheet Tab ID of Conflict History
        TABLE = 0x236, 	        // Data Table
        TABLESTYLE = 0x88F, 	// Table Style
        TABLESTYLEELEMENT = 0x890, 	// Table Style Element
        TABLESTYLES = 0x88E, 	// Table Styles
        TEMPLATE = 0x60,    	// Workbook Is a Template
        THEME = 0x896, 	        // Theme
        TOPMARGIN = 0x28, 	    // Top Margin Measurement
        TXO = 0x1B6, 	        // Text Object
        TXTQUERY = 0x805, 	    //  Text Query Information
        UDDESC = 0xDF, 	        // Description String for Chart Autoformat
        UNCALCED = 0x5E, 	    // Recalculation Status
        USERBVIEW = 0x1A9, 	    // Workbook Custom View Settings
        USERSVIEWBEGIN = 0x1AA, // Custom View Settings
        USERSVIEWEND = 0x1AB, 	// End of Custom View Records
        USESELFS = 0x160, 	    // Natural Language Formulas Flag
        VCENTER = 0x84, 	    // Center Between Vertical Margins
        VERTICALPAGEBREAKS = 0x1A, 	// Explicit Column Page Breaks
        WEBPUB = 0x801, 	    //  Web Publish Item
        WINDOW1 = 0x3D, 	    // Window Information
        WINDOW2 = 0x23E, 	    // Sheet Window Information
        WINDOWPROTECT = 0x19, 	// Windows Are Protected
        WOPT = 0x80B, 	        // Web Options
        WRITEACCESS = 0x5C, 	// Write Access User Name
        WRITEPROT = 0x86, 	    // Workbook Is Write-Protected
        WSBOOL = 0x81, 	        // Additional Workspace Information
        XCT = 0x59, 	        // CRN Record Count
        XF = 0xE0, 	            // Extended Format
        XFCRC = 0x87C, 	        // XF Extensions Checksum
        XFEXT = 0x87D, 	        // XF Extension
        XL5MODIFY = 0x162 	    // Flag for DSF
    }
}
