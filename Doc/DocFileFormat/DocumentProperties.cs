using System;
using System.Collections;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class DocumentProperties : IVisitable
    {
        /// <summary>
        /// True when facing pages should be printed
        /// </summary>
        public bool fFacingPages;

        /// <summary>
        /// True when window control is in effect
        /// </summary>
        public bool fWindowControl;

        /// <summary>
        /// true when doc is a main doc for Print Merge Helper
        /// </summary>
        public bool fPMHMainDoc;

        /// <summary>
        /// Default line suppression storage:<br/>
        /// 0 form letter line supression<br/>
        /// 1 no line supression<br/>
        /// (no longer used)
        /// </summary>
        public short grfSuppression;

        /// <summary>
        /// Footnote position code:<br/>
        /// 0 print as endnotes<br/>
        /// 1 print as bottom of page<br/>
        /// 2 print immediately beneath text
        /// </summary>
        public short Fpc;

        /// <summary>
        /// No longer used
        /// </summary>
        public short grpfIhdt;

        /// <summary>
        /// Restart index for footnotes:<br/>
        /// 0 don't restart note numbering<br/>
        /// 1 restart for each section<br/>
        /// 2 restart for each page
        /// </summary>
        public short rncFtn;

        /// <summary>
        /// Initial footnote number for document
        /// </summary>
        public short nFtn;

        /// <summary>
        /// When true, indicates that information in the hplcpad should 
        /// be refreshed since outline has been dirtied
        /// </summary>
        public bool fOutlineDirtySave;

        /// <summary>
        /// When true, Word believes all pictures recorded in the 
        /// document were created on a Macintosh
        /// </summary>
        public bool fOnlyMacPics;

        /// <summary>
        /// When true, Word believes all pictures recorded in the 
        /// document were created in Windows
        /// </summary>
        public bool fOnlyWinPics;

        /// <summary>
        /// When true, document was created as a print 
        /// merge labels document
        /// </summary>
        public bool fLabelDoc;

        /// <summary>
        /// When true, Word is allowed to hyphenate words that are capitalized
        /// </summary>
        public bool fHyphCapitals;

        /// <summary>
        /// When true, Word will hyphenate newly typed 
        /// text as a background task
        /// </summary>
        public bool fAutoHyphen;

        /// <summary>
        /// 
        /// </summary>
        public bool fFormNoFields;

        /// <summary>
        /// When true, Word will merge styles from its template
        /// </summary>
        public bool fLinkStyles;

        /// <summary>
        /// Whent true, Word will mark revisions as the document is edited
        /// </summary>
        public bool fRevMarking;

        /// <summary>
        /// When true, always make backup when document saved
        /// </summary>
        public bool fBackup;

        /// <summary>
        /// When true, the results of the last Word Count execution are still exactly correct
        /// </summary>
        public bool fExactWords;

        /// <summary>
        /// When true, hidden documents contents are displayed
        /// </summary>
        public bool fPagHidden;

        /// <summary>
        /// When true, field results are displayed, when false, field codes are displayed
        /// </summary>
        public bool fPagResults;

        /// <summary>
        /// When true, annotations are locked for editing
        /// </summary>
        public bool fLockAtn;

        /// <summary>
        /// When true, swap margins on left/right pages
        /// </summary>
        public bool fMirrorMargins;

        /// <summary>
        /// When true, use TrueType fonts by default<br/>
        /// (flag obeyed only when doc was created by WinWord 2.x)
        /// </summary>
        public bool fDflttrueType;

        /// <summary>
        /// When true, file created with SUPPRESSTOPSPACING=YES in Win.ini<br/>
        /// (flag obeyed only when doc was created by WinWord 2.x)
        /// </summary>
        public bool fPagSuppressTopSpacing;

        /// <summary>
        /// When true, document is protected from edit operations
        /// </summary>
        public bool fProtEnabled;

        /// <summary>
        /// When true, restrict selections to occur only within form fields
        /// </summary>
        public bool fDispFormFldSel;

        /// <summary>
        /// When true, show revision markings on screen
        /// </summary>
        public bool fRMView;

        /// <summary>
        /// When true, show revision markings when document is printed
        /// </summary>
        public bool fRMPrint;

        /// <summary>
        /// When true, the current revision marking state is locked
        /// </summary>
        public bool fLockRev;

        /// <summary>
        /// When true, document contains embedded TrueType fonts
        /// </summary>
        public bool fEmbedFonts;

        /// <summary>
        /// Compatibility option: when true, don't add automatic tab 
        /// stops for hanging indent
        /// </summary>
        public bool fNoTabForInd;

        /// <summary>
        /// Compatibility option: when true, don't add extra space 
        /// for raised or lowered characters
        /// </summary>
        public bool fNoSpaceRaiseLower;

        /// <summary>
        /// Compatibility option: when true, suppress the paragraph 
        /// Space Before and Space After options after a page break
        /// </summary>
        public bool fSuppressSpbfAfterPageBreak;

        /// <summary>
        /// Compatibility option: when true, wrap trailing spaces 
        /// at the end of a line to the next line
        /// </summary>
        public bool fWrapTrailSpaces;

        /// <summary>
        /// Compatibility option: when true, print colors as black 
        /// on non-color printer
        /// </summary>
        public bool fMapPrintTextColor;

        /// <summary>
        /// Compatibility option: when true, don't balance columns 
        /// for Continuous Section starts
        /// </summary>
        public bool fNoColumnBalance;

        /// <summary>
        /// 
        /// </summary>
        public bool fConvMailMergeEsc;

        /// <summary>
        /// Compatibility option: when true, suppress extra line 
        /// spacing at top of page
        /// </summary>
        public bool fSuppressTopSpacing;

        /// <summary>
        /// Compatibility option: when true, combine table borders 
        /// like Word 5.x for the Macintosh
        /// </summary>
        public bool fOrigWordTableRules;

        /// <summary>
        /// Compatibility option: when true, don't blank area 
        /// between metafile pictures
        /// </summary>
        public bool fTransparentMetafiles;

        /// <summary>
        /// Compatibility option: when true, show hard page or 
        /// column breaks in frames
        /// </summary>
        public bool fShowBreaksInFrames;

        /// <summary>
        /// Compatibility option: when true, swap left and right 
        /// pages on odd facing pages
        /// </summary>
        public bool fSwapBordersFacingPgs;

        /// <summary>
        /// Default tab width
        /// </summary>
        public ushort dxaTab;

        /// <summary>
        /// Reserved
        /// </summary>
        public ushort wSpare;

        /// <summary>
        /// Width of hyphenation hot zone measured in twips
        /// </summary>
        public ushort dxaHotZ;

        /// <summary>
        /// Number of lines allowed to have consecutive hyphens
        /// </summary>
        public ushort cConsecHypLim;

        /// <summary>
        /// Reserved
        /// </summary>
        public ushort wSpare2;

        /// <summary>
        /// Date and time document was created
        /// </summary>
        public DateAndTime dttmCreated;

        /// <summary>
        /// Date and time document was last revised
        /// </summary>
        public DateAndTime dttmRevised;

        /// <summary>
        /// Date and time document was last printed
        /// </summary>
        public DateAndTime dttmLastPrint;

        /// <summary>
        /// Number of times document has ben revised since its creation
        /// </summary>
        public short nRevision;

        /// <summary>
        /// Time document was last edited
        /// </summary>
        public int tmEdited;

        /// <summary>
        /// Count of words tallied by last Word Count execution
        /// </summary>
        public int cWords;

        /// <summary>
        /// Count of characters tallied by the last Word Count execution
        /// </summary>
        public int cCh;

        /// <summary>
        /// Count of pages tallied by the last Word Count execution
        /// </summary>
        public short cPg;

        /// <summary>
        /// Count of paragraphs tallied by the last Word Count execution
        /// </summary>
        public int cParas;

        /// <summary>
        /// Restart endnote number code:<br/>
        /// 0 don't restart endnote numbering<br/>
        /// 1 restart for each section<br/>
        /// 2 restart for each page
        /// </summary>
        public short rncEdn;

        /// <summary>
        /// Beginning endnote number
        /// </summary>
        public short nEdn;

        /// <summary>
        /// Endnote position code:<br/>
        /// 0 display endnotes at end of section<br/>
        /// 3 display endnotes at the end of document
        /// </summary>
        public short Epc;

        /// <summary>
        /// Number format code for auto footnotes.<br/>
        /// Use the Number Format Table.<br/>
        /// Note: Only the first 16 values in the table can be used.
        /// </summary>
        public short nfcFtnRef;

        /// <summary>
        /// Number format code for auto endnotes.<br/>
        /// Use the Number Format Table.<br/>
        /// Note: Only the first 16 values in the table can be used.
        /// </summary>
        public short nfcEdnRef;

        /// <summary>
        /// Only print data inside of form fields
        /// </summary>
        public bool fPrintFormData;

        /// <summary>
        /// Only save document data that is inside of a form field
        /// </summary>
        public bool fSaveFormData;

        /// <summary>
        /// Shade form fields
        /// </summary>
        public bool fShadeFormData;

        /// <summary>
        /// When true, include footnotes and endnotes in Word Count
        /// </summary>
        public bool fWCFtnEdn;

        /// <summary>
        /// Count of lines tallied by last Word Count operation
        /// </summary>
        public int cLines;

        /// <summary>
        /// Count of words in footnotes and endnotes tallied by last 
        /// word count operation
        /// </summary>
        public int cWordsFtnEdn;

        /// <summary>
        /// Count of characters in footnotes and endnotes tallied by last 
        /// word count operation
        /// </summary>
        public int cChFtnEdn;

        /// <summary>
        /// Count of pages in footnotes and endnotes tallied by last 
        /// word count operation
        /// </summary>
        public int cPgFtnEdn;

        /// <summary>
        /// Count of paragraphs in footnotes and endnotes tallied by last 
        /// word count operation
        /// </summary>
        public int cParasFtnEdn;

        /// <summary>
        /// Count of lines in footnotes and endnotes tallied by last 
        /// word count operation
        /// </summary>
        public int cLinesFtnEdn;

        /// <summary>
        /// Document protection password key only valid if 
        /// fProtEnabled, fLockAtn or fLockRev is true
        /// </summary>
        public int lKeyProtDoc;

        /// <summary>
        /// Document view kind<br/>
        /// 0 Normal view<br/>
        /// 1 Outline view<br/>
        /// 2 Page view
        /// </summary>
        public short wvkSaved;

        /// <summary>
        /// Zoom percentage
        /// </summary>
        public short wScaleSaved;

        /// <summary>
        /// Zoom type:<br/>
        /// 0 None<br/>
        /// 1 Full page<br/>
        /// 2 Page width
        /// </summary>
        public short zkSaved;

        /// <summary>
        /// This is a vertical document<br/>
        /// (Word 6 and 96 only)
        /// </summary>
        public bool fRotateFontW6;

        /// <summary>
        /// Gutter position for this doc:<br/>
        /// 0 Side<br/>
        /// 1 Top
        /// </summary>
        public bool iGutterPos;

        /// <summary>
        /// SUpress extra line spacing at top of page like Word 5.x for the Macintosh
        /// </summary>
        public bool fSuppressTopSpacingMac5;

        /// <summary>
        /// Expand/Codense by whole number of points
        /// </summary>
        public bool fTruncDxaExpand;

        /// <summary>
        /// Print body text before header/footer
        /// </summary>
        public bool fPrintBodyBeforeHdr;

        /// <summary>
        /// Don't add leading (extra space) between rows of text
        /// </summary>
        public bool fNoLeading;

        /// <summary>
        /// USer larger small caps like Word 5.x for the Macintosh
        /// </summary>
        public bool fMWSmallCaps;

        /// <summary>
        /// Autoformat document type: <br/>
        /// 0 for normal<br/>
        /// 1 for letter<br/>
        /// 2 for email
        /// </summary>
        public ushort adt;

        /// <summary>
        /// 
        /// </summary>
        public DocumentTypographyInfo doptypography;

        /// <summary>
        /// 
        /// </summary>
        public DrawingObjectGrid dogrid;

        /// <summary>
        /// Which outline levels are showing in outline view:<br/>
        /// 0 heading 1 only<br/>
        /// 4 headings 1 through 5<br/>
        /// 9 all levels showing
        /// </summary>
        public short lvl;

        /// <summary>
        /// Document has been completely grammar checked
        /// </summary>
        public bool fGramAllDone;

        /// <summary>
        /// No grammar errors exist in document
        /// </summary>
        public bool fGramAllClean;

        /// <summary>
        /// If you are doing font embedding, you should only embed 
        /// the characters in the font that are used in the document
        /// </summary>
        public bool fSubsetFonts;

        /// <summary>
        /// Hide the version created for auto version
        /// </summary>
        public bool fHideLastVersion;

        /// <summary>
        /// This file is based upon an HTML file
        /// </summary>
        public bool fHtmlDoc;

        /// <summary>
        /// Snap table and page borders to page border
        /// </summary>
        public bool fSnapBorder;

        /// <summary>
        /// Place header inside page border
        /// </summary>
        public bool fIncludeHeader;

        /// <summary>
        /// Place footer inside page border
        /// </summary>
        public bool fIncludeFooter;

        /// <summary>
        /// Are we in online view
        /// </summary>
        public bool fForcePageSizePag;

        /// <summary>
        /// Are we auto-promoting fonts to >= hpsZoonFontPag?
        /// </summary>
        public bool fMinFontSizePag;

        /// <summary>
        /// Versioning is turned on
        /// </summary>
        public bool fHaveVersions;

        /// <summary>
        /// Auto versioning is enabled
        /// </summary>
        public bool fAutoVersion;

        /// <summary>
        /// Auto summary info
        /// </summary>
        public AutoSummaryInfo asumyi;

        /// <summary>
        /// Count of characters with spaces
        /// </summary>
        public int cChWS;

        /// <summary>
        /// Count of characters with spaces in footnotes and endnotes
        /// </summary>
        public int cChWSFtnEdn;

        /// <summary>
        /// 
        /// </summary>
        public int grfDocEvents;

        /// <summary>
        /// Have we prompted for virus protection on this document?
        /// </summary>
        public bool fVirusPromted;

        /// <summary>
        /// If prompted, load safely for this document?
        /// </summary>
        public bool fVirusLoadSafe;

        /// <summary>
        /// Random session key to sign above bits for a Word session
        /// </summary>
        public int KeyVirusSession30;

        /// <summary>
        /// Count of double byte characters
        /// </summary>
        public int cDBC;

        /// <summary>
        /// Count of double byte characters in footnotes and endnotes
        /// </summary>
        public int cDBCFtnEdn;

        /// <summary>
        /// Minimum font size if fMinFontSizePag is true
        /// </summary>
        public short hpsZoonFontPag;

        /// <summary>
        /// Height of the window in online view during last repagination
        /// </summary>
        public short dywDispPag;

        /// <summary>
        /// Used internally by Word
        /// </summary>
        public byte ilvlLastBulletMain;

        /// <summary>
        /// Used internally by Word
        /// </summary>
        public byte ilvlLastNumberMain;

        /// <summary>
        /// Default paragraph style for click and type
        /// </summary>
        public short istdClickTypePara;

        /// <summary>
        /// When set to true, language of all text in doc has been auto-detected
        /// </summary>
        public bool fLADAllDone;

        /// <summary>
        /// When set to true, envelope is visible.
        /// </summary>
        public bool fEnvelopeVis;

        /// <summary>
        /// When set to true, doc may have a tentative list in it
        /// </summary>
        public bool fMaybeTentativeListInDoc;

        /// <summary>
        /// When set to 1, doc may have fit text
        /// </summary>
        public bool fMaybeFitText;

        /// <summary>
        /// When set to true, rely on CSS for formatting
        /// </summary>
        public bool fRelyOnCss_WebOpt;

        /// <summary>
        /// When set to true, Rely on VML for displaying graphics in browsers
        /// </summary>
        public bool fRelyOnVML_WebOpt;

        /// <summary>
        /// When set to 1, allow PNG as an output format for graphics
        /// </summary>
        public bool fAllowPNG_WebOpt;

        /// <summary>
        /// Target monitor screen size
        /// </summary>
        public short screenSize_WebOpt;

        /// <summary>
        /// When set to 1, organize supporting files in a folder
        /// </summary>
        public bool fOrganizeInFolder_WebOpt;

        /// <summary>
        /// Use long file names for supporting files
        /// </summary>
        public bool fUseLongFileNames_WebOpt;

        /// <summary>
        /// Target monitor resolution in pixels per inch
        /// </summary>
        public short iPixelsPerInch_WebOpt;

        /// <summary>
        /// When set to 1, the web options have been filled in
        /// </summary>
        public bool fWebOptionsInit;

        /// <summary>
        /// When set to 1, the document may have East Asian layouts
        /// </summary>
        public bool fMaybeFEL;

        /// <summary>
        /// When set to 1, there may be character unit indents or line unit
        /// </summary>
        public bool fCharLineUnits;

        /// <summary>
        /// When set to 1, there may be RTL Tables in this document
        /// </summary>
        public bool fMaybeRTLTables;

        /// <summary>
        /// Compatibility option: when set to true, do not convert 
        /// backslash characters into yen signs
        /// </summary>
        public bool fLeaveBackslashAlone;

        /// <summary>
        /// Compatibility option: when set to true, expand character 
        /// spaces on the line ending SHIFT+RETURN
        /// </summary>
        public bool fExpShRtn;

        /// <summary>
        /// Compatibility option: when set to true, don‘t underline trailing spaces
        /// </summary>
        public bool fDntULTrlSpc;

        /// <summary>
        /// Compatibility option: when set to true, don't balance SBCS and DBCS characters
        /// </summary>
        public bool fDntBlnSbDbWid;

        /// <summary>
        /// Compatibility option: when set to true, add space for underlines.
        /// </summary>
        public bool fMakeSpaceForUL;

        /// <summary>
        /// Compatibility option: suppress extra line spacing like WordPerfect
        /// </summary>
        public bool f2ptExtLeadingOnly;

        /// <summary>
        /// Compatibility option: when set to true, truncate font height
        /// </summary>
        public bool fTruncFontHeight;

        /// <summary>
        /// Compatibility option: when set to true, substitute fonts based on size.
        /// </summary>
        public bool fSubOnSize;

        /// <summary>
        /// Compatibility option: when set to true, lines wrap like Word 6.0
        /// </summary>
        public bool fLineWrapLikeWord6;

        /// <summary>
        /// Compatibility option: when set to true, use Word 6.0/95/97 border rules.
        /// </summary>
        public bool fWW6BorderRules;

        /// <summary>
        /// Compatibility option: when set to true, don't center "exact line height" lines
        /// </summary>
        public bool fExactOnTop;

        /// <summary>
        /// Compatibility option: when set to true, suppress extra line spacing at bottom of page
        /// </summary>
        public bool fExtraAfter;

        /// <summary>
        /// Compatibility option: when set to true, set the width of a space like WordPerfect 5
        /// </summary>
        public bool fWPSpace;

        /// <summary>
        /// Compatibility option: when set to true, do full justification like WordPerfect 6.x
        /// </summary>
        public bool fWPJust;

        /// <summary>
        /// Compatibility option: when set to true, use printer metrics to lay out the document
        /// </summary>
        public bool fPrintMet;

        /// <summary>
        /// Compatibility option: when set to true, lay AutoShapes like Word 97
        /// </summary>
        public bool fSpLayoutLikeWW8;

        /// <summary>
        /// Compatibility option: when set to true, lay footnotes like Word 6.x/95/97.
        /// </summary>
        public bool fFtnLayoutLikeWW8;

        /// <summary>
        /// Compatibility option: when set to true, don't use HTML paragraph auto spacing
        /// </summary>
        public bool fDontUseHTMLParagraphAutoSpacing;

        /// <summary>
        /// Compatibility option: when set to true, don't adjust line height in tables
        /// </summary>
        public bool fDontAdjustLineHeightInTable;

        /// <summary>
        /// Compatibility option: when set to 1, forget last tab alignment
        /// </summary>
        public bool fForgetLastTabAlign;

        /// <summary>
        /// Compatibility option: when set to 1, use auto space like Word 95
        /// </summary>
        public bool fUseAutoSpaceForFullWidthAlpha;

        /// <summary>
        /// Compatibility option: when set to 1, align table rows independently
        /// </summary>
        public bool fAlignTablesRowByRow;

        /// <summary>
        /// Compatibility option: when set to 1, lay out tables with raw width
        /// </summary>
        public bool fLayoutRawTableWidth;

        /// <summary>
        /// Compatibility option: when set to 1, allow table rows to lay out apart
        /// </summary>
        public bool fLayoutTableRowsApart;

        /// <summary>
        /// Compatibility option: when set to 1, use Word 97 line breaking rules for East Asian text
        /// </summary>
        public bool fUserWord97LineBreakingRules;

        /// <summary>
        /// Compatibility option: Do not break wrapped tables across pages.
        /// </summary>
        public bool fDontBreakWrappedTables;

        /// <summary>
        /// Compatibility option: Do not snap text to grid while in a table with inline objects.
        /// </summary>
        public bool fDontSnapToGridInCell;

        /// <summary>
        /// Compatibility option: Select the entire field with the first or last character
        /// </summary>
        public bool fDontAllowFieldEndSelect;

        /// <summary>
        /// Compatibility option: Apply breaking rules
        /// </summary>
        public bool fApplyBreakingRules;

        /// <summary>
        /// Compatibility option: Do not allow hanging punctuation with character grid
        /// </summary>
        public bool fDontWrapTextWithPunct;

        /// <summary>
        /// Compatibility option: Do not use Asian break rules for line breaks with character grid.
        /// </summary>
        public bool fDontUseAsianBreakRules;

        /// <summary>
        /// Compatibility option: Use the Word 2002 table style rules. <br/>
        /// Word 2002 places the top border of a column under the heading row, 
        /// rather than above it as Word 2003 does. <br/>
        /// Word 2003 applies the top border of a column in a more intuitive place when 
        /// there is a header row in the table. This new behavior also fixes an issue with 
        /// shading not displaying correctly for cells using conditional formatting.
        /// </summary>
        public bool fUseWord2002TableStyleRules;

        /// <summary>
        /// Compatibility option:
        /// Allow tables set to ―autofit to contents‖ to extend into the margins when in Print Layout.<br/>
        /// Word 2003 does not allow this by default.
        /// </summary>
        public bool fGrowAutofit;

        /// <summary>
        /// HTML I/O compatibility level
        /// </summary>
        public ushort verCompatPreW10;

        /// <summary>
        /// Page view option
        /// </summary>
        public bool fNoMargPgvwSaved;

        /// <summary>
        /// Page view option
        /// </summary>
        public bool fNoMargPgvWPag;

        /// <summary>
        /// Web View option
        /// </summary>
        public bool fWebViewPag;

        /// <summary>
        /// 
        /// </summary>
        public bool fSeeDrawingsPag;

        /// <summary>
        /// this doc was produced by the document BulletProofer
        /// </summary>
        public bool fBulletProofed;

        /// <summary>
        /// this doc was doctored by the Document Corrupter
        /// </summary>
        public bool fCorrupted;

        /// <summary>
        /// Save option: Embed linguistic in the doc
        /// </summary>
        public bool fSaveUim;

        /// <summary>
        /// Save option: Remove personal information on save
        /// </summary>
        public bool fFilterPrivacy;

        /// <summary>
        /// we are under FReplace (and not just FReplaceRM)
        /// </summary>
        public bool fInFReplaceNoRM;

        /// <summary>
        /// The user has seen the repairs made to the document
        /// </summary>
        public bool fSeenRepairs;

        /// <summary>
        /// XML: The document has XML
        /// </summary>
        public bool fHasXML;

        /// <summary>
        /// 
        /// </summary>
        public bool fSeeScriptAnchorsPag;

        /// <summary>
        /// XML option: Validate XML on save
        /// </summary>
        public bool fValidateXML;

        /// <summary>
        /// XML option: Save the document even if the XML is invalid
        /// </summary>
        public bool fSaveIfInvalidXML;

        /// <summary>
        /// XML option: Show any errors in the XML
        /// </summary>
        public bool fShowXMLErrors;

        /// <summary>
        /// we imported an XML file that had no namespace, so we have elements with no namespace and no schema
        /// </summary>
        public bool fAlwaysMergeEmptyNamespace;

        /// <summary>
        /// 
        /// </summary>
        public int cpMaxListCacheMainDoc;

        /// <summary>
        /// Do not embed system fonts in this document
        /// </summary>
        public bool fDoNotEmbedSystemFont;

        /// <summary>
        /// 
        /// </summary>
        public bool fWordCompact;

        /// <summary>
        /// 
        /// </summary>
        public bool fLiveRecover;

        /// <summary>
        /// Embed smart tags in the document
        /// </summary>
        public bool fEmbedFactoids;

        /// <summary>
        /// Save smart tags as XML properties
        /// </summary>
        public bool fFactoidXML;

        /// <summary>
        /// Done processing smart tags
        /// </summary>
        public bool fFactoidAllDone;

        /// <summary>
        /// Print option: Book fold
        /// </summary>
        public bool fFolioPrint;

        /// <summary>
        /// Print option: Reverse book fold
        /// </summary>
        public bool fReverseFolio;

        /// <summary>
        /// 
        /// </summary>
        public short iTextLineEnding;

        /// <summary>
        /// Do not keep track of formatting
        /// </summary>
        public bool fHideFcc;

        /// <summary>
        /// Track changes: show markup
        /// </summary>
        public bool fAcetateShowMarkup;

        /// <summary>
        /// Track changes: show annotations
        /// </summary>
        public bool fAcetateShowAtn;

        /// <summary>
        /// Track changes: show insertions and deletions
        /// </summary>
        public bool fAcetateShowInsDel;

        /// <summary>
        /// Track changes: show formatting
        /// </summary>
        public bool fAcetateShowProps;

        /// <summary>
        /// Default table style for the document
        /// </summary>
        public ushort istdTableDflt;

        /// <summary>
        /// Internal: Version compatibility for save
        /// </summary>
        public ushort verCompat;

        /// <summary>
        /// Internal: filter state for the Styles and Formatting Pane.
        /// </summary>
        public ushort grfFmtFilter;

        /// <summary>
        /// Book fold printing: sheets per booklet
        /// </summary>
        public short iFolioPages;

        /// <summary>
        /// 
        /// </summary>
        public ushort cpgText;

        /// <summary>
        /// Revision mark CP info
        /// </summary>
        public int cpMinRMText;

        /// <summary>
        /// Revision mark CP info
        /// </summary>
        public int cpMinRMFtn;

        /// <summary>
        /// Revision mark CP info
        /// </summary>
        public int cpMinRMHdd;

        /// <summary>
        /// Revision mark CP info
        /// </summary>
        public int cpMinRMAtn;

        /// <summary>
        /// Revision mark CP info
        /// </summary>
        public int cpMinRMEdn;

        /// <summary>
        /// Revision mark CP info
        /// </summary>
        public int cpMinRMTxbx;

        /// <summary>
        /// Revision mark CP info
        /// </summary>
        public int cpMinRMHdrTxbx;

        /// <summary>
        /// 
        /// </summary>
        public int rsidRoot;

        /// <summary>
        /// Document Protection: Treat lock for annotations as Read Only
        /// </summary>
        public bool fTreatLockAtnAsReadOnly;

        /// <summary>
        /// Document Protection: Style lockdown is turned on
        /// </summary>
        public bool fStyleLock;

        /// <summary>
        /// Document Protection: Allow AutoFormat to override style lockdown
        /// </summary>
        public bool fAutoFmtOverride;

        /// <summary>
        /// XML Option: Remove Word XML when saving; save only non-Word XML data.
        /// </summary>
        public bool fRemoveWordML;

        /// <summary>
        /// XML Option: Apply custom transform on Save
        /// </summary>
        public bool fApplyCustomXForm;

        /// <summary>
        /// Document Protection: Style lockdown is enforced
        /// </summary>
        public bool fStyeLockEnforced;

        /// <summary>
        /// Document Protection: Simulate locked for annotations in 
        /// older version when a document has style protection
        /// </summary>
        public bool fFakeLockAtn;

        /// <summary>
        /// XML Option: Ignore mixed content
        /// </summary>
        public bool fIgnoreMixedContent;

        /// <summary>
        /// XML Option: Show placeholder text for all empty XML elements
        /// </summary>
        public bool fShowPlaceholderText;

        /// <summary>
        /// 
        /// </summary>
        public uint grf;

        /// <summary>
        /// Reading mode: ink lock down
        /// </summary>
        public bool fReadingModeInkLockDown;

        /// <summary>
        /// Track changes: Show ink annotations
        /// </summary>
        public bool fAcetateShowInkAtn;

        /// <summary>
        /// Filter date and time
        /// </summary>
        public bool fFilterDttm;

        /// <summary>
        /// Enforce document protection
        /// </summary>
        public bool fEnforceDocProt;

        /// <summary>
        /// Doc protection level:<br/>
        /// 0 Protect for track changes<br/>
        /// 1 Comment protection<br/>
        /// 2 Form protection<br/>
        /// 3 Read Only
        /// </summary>
        public ushort iDocProtCur;

        /// <summary>
        /// 
        /// </summary>
        public bool fDispBkSpSaved;

        /// <summary>
        /// Reading Layout page size lockdown
        /// </summary>
        public short dxaPageLock;

        /// <summary>
        /// Reading Layout page size lockdown
        /// </summary>
        public short dyaPageLock;

        /// <summary>
        /// Reading Layout font lockdown
        /// </summary>
        public int pctFontLock;

        /// <summary>
        /// 
        /// </summary>
        public byte grfitbid;

        /// <summary>
        /// Number of LFOs when CleanupLists last attempted cleaning
        /// </summary>
        public ushort ilfoMacAtCleanup;

        /// <summary>
        /// Parses the bytes to retrieve a DocumentProperties
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public DocumentProperties(FileInformationBlock fib, VirtualStream tableStream)
        {
            setDefaultCompatibilityOptions(fib.nFib);

            var bytes = new byte[fib.lcbDop];
            tableStream.Read(bytes, 0, (int)fib.lcbDop, fib.fcDop);

            try
            {
                if (bytes.Length > 0)
                {
                    BitArray bits;

                    //split byte 0 and 1 into bits
                    bits = new BitArray(new byte[] { bytes[0], bytes[1] });
                    this.fFacingPages = bits[0];
                    this.fWindowControl = bits[1];
                    this.fPMHMainDoc = bits[2];
                    this.grfSuppression = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 3, 2));
                    this.Fpc = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 5, 2));

                    //split byte 2 and 3 into bits
                    bits = new BitArray(new byte[] { bytes[2], bytes[3] });
                    this.rncFtn = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 0, 2));
                    this.nFtn = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 2, 14));

                    //split byte 4 and 5 into bits
                    bits = new BitArray(new byte[] { bytes[4], bytes[5] });
                    this.fOutlineDirtySave = bits[0];
                    this.fOnlyMacPics = bits[8];
                    this.fOnlyWinPics = bits[9];
                    this.fLabelDoc = bits[10];
                    this.fHyphCapitals = bits[11];
                    this.fAutoHyphen = bits[12];
                    this.fFormNoFields = bits[13];
                    this.fLinkStyles = bits[14];
                    this.fRevMarking = bits[15];

                    //split byte 6 and 7 into bits
                    bits = new BitArray(new byte[] { bytes[6], bytes[7] });
                    this.fBackup = bits[0];
                    this.fExactWords = bits[1];
                    this.fPagHidden = bits[2];
                    this.fPagResults = bits[3];
                    this.fLockAtn = bits[4];
                    this.fMirrorMargins = bits[5];
                    //bit 6 is reserved
                    this.fDflttrueType = bits[7];
                    this.fProtEnabled = bits[8];
                    this.fDispFormFldSel = bits[9];
                    this.fRMView = bits[10];
                    this.fRMPrint = bits[11];
                    //bit 12 and 13 are reserved
                    this.fLockRev = bits[14];
                    this.fEmbedFonts = bits[15];

                    //split byte 8 and 9 into bits
                    bits = new BitArray(new byte[] { bytes[8], bytes[9] });
                    this.fNoTabForInd = bits[0];
                    this.fNoSpaceRaiseLower = bits[1];
                    this.fSuppressSpbfAfterPageBreak = bits[2];
                    this.fWrapTrailSpaces = bits[3];
                    this.fMapPrintTextColor = bits[4];
                    this.fNoColumnBalance = bits[5];
                    this.fConvMailMergeEsc = bits[6];
                    this.fSuppressTopSpacing = bits[7];
                    this.fOrigWordTableRules = bits[8];
                    this.fTransparentMetafiles = bits[9];
                    this.fShowBreaksInFrames = bits[10];
                    this.fSwapBordersFacingPgs = bits[11];

                    this.dxaTab = System.BitConverter.ToUInt16(bytes, 10);
                    this.dxaHotZ = System.BitConverter.ToUInt16(bytes, 14);
                    this.cConsecHypLim = System.BitConverter.ToUInt16(bytes, 16);

                    var createdbytes = new byte[4];
                    Array.Copy(bytes, 20, createdbytes, 0, createdbytes.Length);
                    this.dttmCreated = new DateAndTime(createdbytes);

                    var revisedbytes = new byte[4];
                    Array.Copy(bytes, 24, revisedbytes, 0, revisedbytes.Length);
                    this.dttmRevised = new DateAndTime(revisedbytes);

                    var printbytes = new byte[4];
                    Array.Copy(bytes, 28, printbytes, 0, printbytes.Length);
                    this.dttmLastPrint = new DateAndTime(printbytes);

                    this.nRevision = System.BitConverter.ToInt16(bytes, 32);
                    this.tmEdited = System.BitConverter.ToInt32(bytes, 34);
                    this.cWords = System.BitConverter.ToInt32(bytes, 38);
                    this.cCh = System.BitConverter.ToInt32(bytes, 42);
                    this.cPg = System.BitConverter.ToInt16(bytes, 46);
                    this.cParas = System.BitConverter.ToInt32(bytes, 48);

                    //split byte 52 and 53 into bits
                    bits = new BitArray(new byte[] { bytes[52], bytes[53] });
                    this.rncEdn = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 0, 2));
                    this.nEdn = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 2, 14));

                    //split byte 54 and 55 into bits
                    bits = new BitArray(new byte[] { bytes[54], bytes[55] });
                    this.Epc = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 0, 2));
                    this.nfcFtnRef = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 2, 4));
                    this.nfcEdnRef = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 6, 4));
                    this.fPrintFormData = bits[10];
                    this.fSaveFormData = bits[11];
                    this.fShadeFormData = bits[12];
                    //bits 13 and 14 are reserved
                    this.fWCFtnEdn = bits[15];

                    this.cLines = System.BitConverter.ToInt32(bytes, 56);
                    this.cWordsFtnEdn = System.BitConverter.ToInt32(bytes, 60);
                    this.cChFtnEdn = System.BitConverter.ToInt32(bytes, 64);
                    this.cPgFtnEdn = System.BitConverter.ToInt16(bytes, 68);
                    this.cParasFtnEdn = System.BitConverter.ToInt32(bytes, 70);
                    this.cLinesFtnEdn = System.BitConverter.ToInt32(bytes, 74);
                    this.lKeyProtDoc = System.BitConverter.ToInt32(bytes, 78);

                    //split byte 82 and 83 into bits
                    bits = new BitArray(new byte[] { bytes[82], bytes[83] });
                    this.wvkSaved = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 0, 3));
                    this.wScaleSaved = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 3, 9));
                    this.zkSaved = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 12, 2));
                    this.fRotateFontW6 = bits[14];
                    this.iGutterPos = bits[15];

                    //compatibility options section
                    if (bytes.Length > 84)
                    {
                        //split byte 84,85,86,87 into bits
                        bits = new BitArray(new byte[] { bytes[84], bytes[85], bytes[86], bytes[87] });
                        this.fNoTabForInd = bits[0];
                        this.fNoSpaceRaiseLower = bits[1];
                        this.fSuppressSpbfAfterPageBreak = bits[2];
                        this.fWrapTrailSpaces = bits[3];
                        this.fMapPrintTextColor = bits[4];
                        this.fNoColumnBalance = bits[5];
                        this.fConvMailMergeEsc = bits[6];
                        this.fSuppressTopSpacing = bits[7];
                        this.fOrigWordTableRules = bits[8];
                        this.fTransparentMetafiles = bits[9];
                        this.fShowBreaksInFrames = bits[10];
                        this.fSwapBordersFacingPgs = bits[11];
                        //bits 12,13,14,15 are reserved
                        this.fSuppressTopSpacingMac5 = bits[16];
                        this.fTruncDxaExpand = bits[17];
                        this.fPrintBodyBeforeHdr = bits[18];
                        this.fNoLeading = bits[19];
                        //bits 20 is reserved
                        this.fMWSmallCaps = bits[21];
                        //bits 22-31 are reserved

                        if (bytes.Length > 88)
                        {
                            this.adt = (ushort)System.BitConverter.ToInt16(bytes, 88);

                            var doptypoBytes = new byte[310];
                            Array.Copy(bytes, 90, doptypoBytes, 0, doptypoBytes.Length);
                            this.doptypography = new DocumentTypographyInfo(doptypoBytes);

                            var dogridBytes = new byte[10];
                            Array.Copy(bytes, 400, dogridBytes, 0, dogridBytes.Length);
                            this.dogrid = new DrawingObjectGrid(dogridBytes);

                            //split byte 410 and 411 into bits
                            bits = new BitArray(new byte[] { bytes[410], bytes[411] });
                            //bit 0 is reserved
                            this.lvl = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 1, 4));
                            this.fGramAllDone = bits[5];
                            this.fGramAllClean = bits[6];
                            this.fSubsetFonts = bits[7];
                            this.fHideLastVersion = bits[8];
                            this.fHtmlDoc = bits[9];
                            //bit 10 is reserved
                            this.fSnapBorder = bits[11];
                            this.fIncludeHeader = bits[12];
                            this.fIncludeFooter = bits[13];
                            this.fForcePageSizePag = bits[14];
                            this.fMinFontSizePag = bits[15];

                            //split byte 412 and 413 into bits
                            bits = new BitArray(new byte[] { bytes[412], bytes[413] });
                            this.fHaveVersions = bits[0];
                            this.fAutoVersion = bits[1];
                            //other bits are reserved

                            var asumybits = new byte[12];
                            Array.Copy(bytes, 414, asumybits, 0, asumybits.Length);
                            this.asumyi = new AutoSummaryInfo(asumybits);

                            this.cChWS = System.BitConverter.ToInt32(bytes, 426);
                            this.cChWSFtnEdn = System.BitConverter.ToInt32(bytes, 430);
                            this.grfDocEvents = System.BitConverter.ToInt32(bytes, 434);

                            //split bytes 438-441 in bits
                            bits = new BitArray(new byte[] { bytes[438], bytes[439], bytes[440], bytes[441] });
                            this.fVirusPromted = bits[0];
                            this.fVirusLoadSafe = bits[1];
                            this.KeyVirusSession30 = (int)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 2, 30));

                            this.cDBC = System.BitConverter.ToInt32(bytes, 480);
                            this.cDBCFtnEdn = System.BitConverter.ToInt32(bytes, 484);
                            this.nfcEdnRef = System.BitConverter.ToInt16(bytes, 492);
                            this.nfcFtnRef = System.BitConverter.ToInt16(bytes, 494);
                            this.hpsZoonFontPag = System.BitConverter.ToInt16(bytes, 496);
                            this.dywDispPag = System.BitConverter.ToInt16(bytes, 498);

                            //WORD 2000, 2002, 2003 PART
                            if (bytes.Length > 500)
                            {
                                this.ilvlLastBulletMain = bytes[500];
                                this.ilvlLastNumberMain = bytes[501];
                                this.istdClickTypePara = System.BitConverter.ToInt16(bytes, 502);

                                //split byte 504 and 505 into bits
                                bits = new BitArray(new byte[] { bytes[504], bytes[505] });
                                this.fLADAllDone = bits[0];
                                this.fEnvelopeVis = bits[1];
                                this.fMaybeTentativeListInDoc = bits[2];
                                this.fMaybeFitText = bits[3];
                                this.fRelyOnCss_WebOpt = bits[9];
                                this.fRelyOnVML_WebOpt = bits[10];
                                this.fAllowPNG_WebOpt = bits[11];
                                this.screenSize_WebOpt = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 12, 4));

                                //split byte 506 and 507 into bits
                                bits = new BitArray(new byte[] { bytes[506], bytes[507] });
                                this.fOrganizeInFolder_WebOpt = bits[0];
                                this.fUseLongFileNames_WebOpt = bits[1];
                                this.iPixelsPerInch_WebOpt = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 2, 10));
                                this.fWebOptionsInit = bits[12];
                                this.fMaybeFEL = bits[13];
                                this.fCharLineUnits = bits[14];
                                this.fMaybeRTLTables = bits[15];

                                //split bytes 508,509,510,511 into bits
                                bits = new BitArray(new byte[] { bytes[508], bytes[509], bytes[510], bytes[511] });
                                this.fNoTabForInd = bits[0];
                                this.fNoSpaceRaiseLower = bits[1];
                                this.fSuppressSpbfAfterPageBreak = bits[2];
                                this.fWrapTrailSpaces = bits[3];
                                this.fMapPrintTextColor = bits[4];
                                this.fNoColumnBalance = bits[5];
                                this.fConvMailMergeEsc = bits[6];
                                this.fSuppressTopSpacing = bits[7];
                                this.fOrigWordTableRules = bits[8];
                                this.fTransparentMetafiles = bits[9];
                                this.fShowBreaksInFrames = bits[10];
                                this.fSwapBordersFacingPgs = bits[11];
                                this.fLeaveBackslashAlone = bits[12];
                                this.fExpShRtn = bits[13];
                                this.fDntULTrlSpc = bits[14];
                                this.fDntBlnSbDbWid = bits[15];
                                this.fSuppressTopSpacingMac5 = bits[16];
                                this.fTruncDxaExpand = bits[17];
                                this.fPrintBodyBeforeHdr = bits[18];
                                this.fNoLeading = bits[19];
                                this.fMakeSpaceForUL = bits[20];
                                this.fMWSmallCaps = bits[21];
                                this.f2ptExtLeadingOnly = bits[22];
                                this.fTruncFontHeight = bits[23];
                                this.fSubOnSize = bits[24];
                                this.fLineWrapLikeWord6 = bits[25];
                                this.fWW6BorderRules = bits[26];
                                this.fExactOnTop = bits[27];
                                this.fExtraAfter = bits[28];
                                this.fWPSpace = bits[29];
                                this.fWPJust = bits[30];
                                this.fPrintMet = bits[31];

                                //split bytes 512,513,514,515 into bits
                                bits = new BitArray(new byte[] { bytes[512], bytes[513], bytes[514], bytes[515] });
                                this.fSpLayoutLikeWW8 = bits[0];
                                this.fFtnLayoutLikeWW8 = bits[1];
                                this.fDontUseHTMLParagraphAutoSpacing = bits[2];
                                this.fDontAdjustLineHeightInTable = bits[3];
                                this.fForgetLastTabAlign = bits[4];
                                this.fUseAutoSpaceForFullWidthAlpha = bits[5];
                                this.fAlignTablesRowByRow = bits[6];
                                this.fLayoutRawTableWidth = bits[7];
                                this.fLayoutTableRowsApart = bits[8];
                                this.fUserWord97LineBreakingRules = bits[9];
                                this.fDontBreakWrappedTables = bits[10];
                                this.fDontSnapToGridInCell = bits[11];
                                this.fDontAllowFieldEndSelect = bits[12];
                                this.fApplyBreakingRules = bits[13];
                                this.fDontWrapTextWithPunct = bits[14];
                                this.fDontUseAsianBreakRules = bits[15];
                                this.fUseWord2002TableStyleRules = bits[16];
                                this.fGrowAutofit = bits[17];
                                //bits 18-31 are unused

                                //bytes 516-539 are unused

                                //split bytes 540,541,542,543 into bits
                                bits = new BitArray(new byte[] { bytes[540], bytes[541], bytes[542], bytes[543] });
                                this.verCompatPreW10 = (ushort)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 0, 16));
                                this.fNoMargPgvwSaved = bits[16];
                                this.fNoMargPgvWPag = bits[17];
                                this.fWebViewPag = bits[18];
                                this.fSeeDrawingsPag = bits[19];
                                this.fBulletProofed = bits[20];
                                this.fCorrupted = bits[21];
                                this.fSaveUim = bits[22];
                                this.fFilterPrivacy = bits[23];
                                this.fInFReplaceNoRM = bits[24];
                                this.fSeenRepairs = bits[25];
                                this.fHasXML = bits[26];
                                this.fSeeScriptAnchorsPag = bits[27];
                                this.fValidateXML = bits[28];
                                this.fSaveIfInvalidXML = bits[29];
                                this.fShowXMLErrors = bits[30];
                                this.fAlwaysMergeEmptyNamespace = bits[31];

                                this.cpMaxListCacheMainDoc = System.BitConverter.ToInt32(bytes, 544);

                                //split bytes 548,549 into bits
                                bits = new BitArray(new byte[] { bytes[548], bytes[549] });
                                this.fDoNotEmbedSystemFont = bits[0];
                                this.fWordCompact = bits[1];
                                this.fLiveRecover = bits[2];
                                this.fEmbedFactoids = bits[3];
                                this.fFactoidXML = bits[4];
                                this.fFactoidAllDone = bits[5];
                                this.fFolioPrint = bits[6];
                                this.fReverseFolio = bits[7];
                                this.iTextLineEnding = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 8, 3));
                                this.fHideFcc = bits[11];
                                this.fAcetateShowMarkup = bits[12];
                                this.fAcetateShowAtn = bits[13];
                                this.fAcetateShowInsDel = bits[14];
                                this.fAcetateShowProps = bits[15];

                                this.istdTableDflt = System.BitConverter.ToUInt16(bytes, 550);
                                this.verCompat = System.BitConverter.ToUInt16(bytes, 552);
                                this.grfFmtFilter = System.BitConverter.ToUInt16(bytes, 554);
                                this.iFolioPages = System.BitConverter.ToInt16(bytes, 556);
                                this.cpgText = System.BitConverter.ToUInt16(bytes, 558);
                                this.cpMinRMText = System.BitConverter.ToInt32(bytes, 560);
                                this.cpMinRMFtn = System.BitConverter.ToInt32(bytes, 564);
                                this.cpMinRMHdd = System.BitConverter.ToInt32(bytes, 568);
                                this.cpMinRMAtn = System.BitConverter.ToInt32(bytes, 572);
                                this.cpMinRMEdn = System.BitConverter.ToInt32(bytes, 576);
                                this.cpMinRMTxbx = System.BitConverter.ToInt32(bytes, 580);
                                this.cpMinRMHdrTxbx = System.BitConverter.ToInt32(bytes, 584);
                                this.rsidRoot = System.BitConverter.ToInt32(bytes, 588);

                                if (bytes.Length == 610)
                                {
                                    //split bytes 592,593,594,595 into bits
                                    bits = new BitArray(new byte[] { bytes[592], bytes[593], bytes[594], bytes[595] });
                                    this.fTreatLockAtnAsReadOnly = bits[0];
                                    this.fStyleLock = bits[1];
                                    this.fAutoFmtOverride = bits[2];
                                    this.fRemoveWordML = bits[3];
                                    this.fApplyCustomXForm = bits[4];
                                    this.fStyeLockEnforced = bits[5];
                                    this.fFakeLockAtn = bits[6];
                                    this.fIgnoreMixedContent = bits[7];
                                    this.fShowPlaceholderText = bits[8];
                                    this.grf = Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 9, 23));

                                    //split bytes 596 and 597 into bits
                                    bits = new BitArray(new byte[] { bytes[596], bytes[597] });
                                    this.fReadingModeInkLockDown = bits[0];
                                    this.fAcetateShowInkAtn = bits[1];
                                    this.fFilterDttm = bits[2];
                                    this.fEnforceDocProt = bits[3];
                                    this.iDocProtCur = (ushort)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 4, 3));
                                    this.fDispBkSpSaved = bits[7];

                                    this.dxaPageLock = System.BitConverter.ToInt16(bytes, 598);
                                    this.dyaPageLock = System.BitConverter.ToInt16(bytes, 600);
                                    this.pctFontLock = System.BitConverter.ToInt32(bytes, 602);
                                    this.grfitbid = bytes[606];
                                    //byte 607 is unused
                                    this.ilfoMacAtCleanup = System.BitConverter.ToUInt16(bytes, 608);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                //this DOP was probably not written by Word
                TraceLogger.Warning("Unexpected length of DOP ({0} bytes) in input file.", (int)fib.lcbDop);
            }
        }

        private void setDefaultCompatibilityOptions(FileInformationBlock.FibVersion nFib)
        {
            if(nFib == FileInformationBlock.FibVersion.Fib1997 || nFib == FileInformationBlock.FibVersion.Fib1997Beta)
            {
                //Word 97 default settings
                this.fAlignTablesRowByRow = true;
                this.fLayoutTableRowsApart = true;
                this.fGrowAutofit = true;
                this.fDontWrapTextWithPunct = true;
                //ToDo: Don't autofit tables next to wrapped objects
                //ToDo: Don't break constrained tables forced onto the page
                this.fDontBreakWrappedTables = true;
                this.fDontSnapToGridInCell = true;
                this.fDontUseAsianBreakRules = true;
                this.fNoTabForInd = true;
                this.fDontUseHTMLParagraphAutoSpacing = true;
                this.fForgetLastTabAlign = true;
                this.fSpLayoutLikeWW8 = true;
                this.fFtnLayoutLikeWW8 = true;
                this.fLayoutRawTableWidth = true;
                this.fDontAllowFieldEndSelect = true;
                //ToDo: underline characters in numbered lists
                this.fUseWord2002TableStyleRules = true;
                this.fUserWord97LineBreakingRules = true;
            }
            else if (nFib == FileInformationBlock.FibVersion.Fib2000)
            {
                //Word 2000 default settings

                this.fGrowAutofit = true;
                this.fDontWrapTextWithPunct = true;
                //ToDo: Don't autofit tables next to wrapped objects
                this.fDontBreakWrappedTables = true;
                this.fDontSnapToGridInCell = true;
                this.fDontUseAsianBreakRules = true;
                this.fNoTabForInd = true;
                this.fDontAllowFieldEndSelect = true;
                //ToDo: underline characters in numbered lists
                this.fUseWord2002TableStyleRules = true;
            }
            else if(nFib ==  FileInformationBlock.FibVersion.Fib2002)
            {
                //Word 2002 (XP)

                this.fGrowAutofit = true;
                //ToDo: Don't autofit tables next to wrapped objects
                this.fDontBreakWrappedTables = true;
                this.fNoTabForInd = true;
                this.fUseWord2002TableStyleRules = true;
            }
            else if (nFib == FileInformationBlock.FibVersion.Fib2003)
            {
                //Word 2003

                //ToDo: Don't autofit tables next to wrapped objects
                this.fDontBreakWrappedTables = true;
                this.fNoTabForInd = true;

            }
            else if (nFib < FileInformationBlock.FibVersion.Fib1997Beta)
            {
                throw new UnspportedFileVersionException();
            }
        }

        #region IVisitable Members

        public virtual void Convert<T>(T mapping)
        {
            ((IMapping<DocumentProperties>)mapping).Apply(this);
        }

        #endregion
    }
}
