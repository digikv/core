/*
 *
 * XLSX Class - Harbour XLSX library source code:
 *
 * Copyright 2018 Srdjan Dragojlovic <digikv@yahoo.com> - S.O.K. doo Kraljevo
 * www - http://www.sokdoo.com
*/

#require "hbmzip"
#include "simpleio.ch"
#INCLUDE "hbclass.ch"

REQUEST HB_CODEPAGE_UTF8EX

STATIC aHorizontal := {"left","center","right"}
STATIC aVertical := {"top","center","bottom"}
STATIC aBordersStyle := { "thin", "medium", "thick", "double", "dashed", "dotted", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot" }
STATIC aPattern := {"solid", "horzStripe", "vertStripe", "reverseDiagStripe", "diagStripe", "horzCross", "diagCross", "thinHorzStripe", "thinVertStripe", "thinReverseDiagStripe", "thinDiagStripe", "thinHorzCross", "thinDiagCross" }

CLASS WorkBook
DATA cTempDir PROTECTED
DATA cName
DATA Company INIT "S.O.K. doo Kraljevo, Serbia"
DATA Application INIT "XLSX Class"
DATA DocSecurity INIT 0
DATA ScaleCrop INIT .F.
DATA LinksUpToDate INIT .F.
DATA SharedDoc INIT .F.
DATA HyperlinksChanged INIT .F.
DATA AppVersion INIT "12.0000"
DATA aWorkSheetNames PROTECTED  
DATA aWorkSheetObjects PROTECTED 
DATA aWorkbookFonts PROTECTED
DATA aWorkbookStyles PROTECTED
DATA aWorkbookFills PROTECTED
DATA aWorkbookBorders PROTECTED
DATA aWorkbookDrawings PROTECTED
DATA aNumFormat PROTECTED
DATA aSharedStrings INIT {}
DATA _str_total INIT 0
METHOD NewFont( cFont, nFontSize, lBold, lItalic, lUnderline, lStrike, cRGB )
METHOD NewFillPattern( nFillPattern, cFG, cBG )
METHOD NewBorder( nTL, nTR, nTT, nTB, nTD, cCL, cCR, cCT, cCB, cCD  )
METHOD NewStyle( nFont, nBorder, nFill, nVA, nHA, nRotation, lWrap )
METHOD NewFormat( cFormat )
METHOD New(cName)
METHOD WorkSheet(cName)
METHOD Save()
ENDCLASS

METHOD NewFillPattern( nFillPattern, cFG, cBG ) CLASS WorkBook
LOCAL nPos := 0, c:= ""
IF HB_ISNUMERIC(nFillPattern) .AND. nFillPattern>0 .AND. nFillPattern<14 .AND. HB_ISSTRING( cBG ) .AND. HB_ISSTRING( cFG )
	c := ALLTRIM(STR(nFillPattern,2,0))+","+cFG+","+cBG
	IF (nPos:=ASCAN( ::aWorkbookFills, c ) )==0
		AADD( ::aWorkbookFills, c )
		nPos := LEN( ::aWorkbookFills )
	ENDIF
ENDIF
Return nPos

METHOD NewBorder( nTL, nTR, nTT, nTB, nTD, cCL, cCR, cCT, cCB, cCD  ) CLASS WorkBook
LOCAL nPos := 0
LOCAL c := ""
c += IIF( (HB_ISNUMERIC(nTL) .AND. nTL>0 .AND. nTL<14), ALLTRIM(STR(nTL,2,0))+IIF( (!HB_ISNIL(cCL) .AND. HB_ISSTRING( cCL )), ","+cCL+",", ",000000," ),"0,000000," )
c += IIF( (HB_ISNUMERIC(nTR) .AND. nTR>0 .AND. nTR<14), ALLTRIM(STR(nTR,2,0))+IIF( (!HB_ISNIL(cCR) .AND. HB_ISSTRING( cCR )), ","+cCR+",", ",000000," ),"0,000000," )
c += IIF( (HB_ISNUMERIC(nTT) .AND. nTT>0 .AND. nTT<14), ALLTRIM(STR(nTT,2,0))+IIF( (!HB_ISNIL(cCT) .AND. HB_ISSTRING( cCT )), ","+cCT+",", ",000000," ),"0,000000," )
c += IIF( (HB_ISNUMERIC(nTB) .AND. nTB>0 .AND. nTB<14), ALLTRIM(STR(nTB,2,0))+IIF( (!HB_ISNIL(cCB) .AND. HB_ISSTRING( cCB )), ","+cCB+",", ",000000," ),"0,000000," )
c += IIF( (HB_ISNUMERIC(nTD) .AND. nTD>0 .AND. nTD<14), ALLTRIM(STR(nTD,2,0))+IIF( (!HB_ISNIL(cCD) .AND. HB_ISSTRING( cCD )), ","+cCD+",", ",000000," ),"0,000000" )
IF (nPos:=ASCAN( ::aWorkbookBorders, c ) )==0
	AADD( ::aWorkbookBorders, c )
	nPos := LEN( ::aWorkbookBorders )
ENDIF
Return nPos

METHOD NewFormat( cFormat ) CLASS WorkBook
Local nPos := 0
IF HB_ISSTRING( cFormat )
	IF (nPos:=ASCAN( ::aNumFormat, cFormat ) )==0
		AADD( ::aNumFormat, cFormat )
		nPos := LEN( ::aNumFormat )
	ENDIF
ENDIF
Return nPos

METHOD NewFont( cFont, nFontSize, lBold, lItalic, lUnderline, lStrike, cRGB ) CLASS WorkBook
Local nPos := 0
cFont := iif( HB_ISNIL( cFont ), "Calibri", cFont )
cFont += ","+iif( HB_ISNIL( nFontSize ) .OR. !HB_ISNUMERIC( nFontSize ), "11", ALLTRIM(STR(nFontSize,10,0)) )
cFont += ","+iif( HB_ISNIL( lBold ) .OR. !HB_ISLOGICAL( lBold ), "0", iif(lBold,"1","0") )
cFont += ","+iif( HB_ISNIL( lItalic ) .OR. !HB_ISLOGICAL( lItalic ), "0", iif(lItalic,"1","0") ) 
cFont += ","+iif( HB_ISNIL( lUnderline ) .OR. !HB_ISLOGICAL( lUnderline ), "0", iif(lUnderline,"1","0") )
cFont += ","+iif( HB_ISNIL( lStrike ) .OR. !HB_ISLOGICAL( lStrike ), "0", iif(lStrike,"1","0") )
cFont += ","+iif( HB_ISNIL( cRGB ) .OR. !HB_ISSTRING( cRGB ), "000000", cRGB )
IF (nPos:=ASCAN( ::aWorkbookFonts, cFont ) )==0
	AADD( ::aWorkbookFonts, cFont )
	nPos := LEN( ::aWorkbookFonts )
ENDIF
Return nPos

METHOD NewStyle( nFont, nBorder, nFill, nVA, nHA, nFormat, nRotation, lWrap ) CLASS WorkBook
Local nPos := 0, c := ""
c := iif( HB_ISNIL( nFont ) .OR. !HB_ISNUMERIC( nFont ), "0", ALLTRIM(STR(nFont-1,10,0)) )
c += ","+iif( HB_ISNIL( nBorder ) .OR. !HB_ISNUMERIC( nBorder ), "0", ALLTRIM(STR(nBorder-1,10,0)) )
c += ","+iif( HB_ISNIL( nFill ) .OR. !HB_ISNUMERIC( nFill ), "0", ALLTRIM(STR(nFill-1,10,0)) )
c += ","+iif( HB_ISNIL( nVA ) .OR. !HB_ISNUMERIC( nVA ) .OR. nVA<1 .OR. nVa>3, "0", ALLTRIM(STR(nVA,10,0)) )
c += ","+iif( HB_ISNIL( nHA ) .OR. !HB_ISNUMERIC( nHA ) .OR. nHA<1 .OR. nHa>3, "0", ALLTRIM(STR(nHA,10,0)) )
c += ","+iif( HB_ISNIL( nFormat ) .OR. !HB_ISNUMERIC( nFormat ), "0", ALLTRIM(STR(nFormat+163,10,0)) )
c += ","+iif( HB_ISNIL( nRotation ) .OR. !HB_ISNUMERIC( nRotation ), "0", ALLTRIM(STR(nRotation,10,0)) )
c += ","+iif( HB_ISNIL( lWrap ) .OR. !HB_ISLOGICAL( lWrap ), "0", "1" )
IF (nPos:=ASCAN( ::aWorkbookStyles, c ) )==0
	AADD( ::aWorkbookStyles, c )
	nPos := LEN( ::aWorkbookStyles )
ENDIF
Return nPos+1

METHOD New(cName) CLASS WorkBook
::cName := cName
::aWorkSheetNames := {}  
::aWorkSheetObjects := {}
::aWorkbookDrawings := {}
::aWorkbookFonts := { "Calibri,11,0,0,0,0,000000" }
::aWorkbookStyles := {}
::aWorkbookFills := { "0,00000000,00000000", "0,00000000,00000000" }
::aWorkbookBorders := { "0,0,0,0,0" }
::aNumFormat := { "General", Set( _SET_DATEFORMAT ) }
Return Self

METHOD WorkSheet( cName ) CLASS WorkBook
LOCAL oWorkSheet, nPos 
IF (nPos:=ASCAN(::aWorkSheetNames, cName))==0
    oWorkSheet:= WorkSheet():New( cName )
	oWorkSheet:oParent := Self
	aADD( ::aWorkSheetNames, cName )
	aADD( ::aWorkSheetObjects, oWorkSheet )
	nPos := LEN( ::aWorkSheetNames )
ENDIF
Return ::aWorkSheetObjects[nPos]

METHOD Save() CLASS WorkBook
LOCAL n := hb_RandomInt( 1, 10000 )
LOCAL cSep := hb_ps(), i, handle, c
LOCAL cDate := DTOS( DATE() ), cTime := TIME(), cDateTime := LEFT( cDate, 4 )+'-'+SUBSTR(cDate,5,2)+'-'+RIGHT(cDate,2)+'T'+LEFT(cTime,8)
::cTempDir := hb_DirSepToOS( hb_DirTemp()+"ExcelTemp" )

hb_DirRemoveAll( ::cTempDir )
IF MakeDir( ::cTempDir ) <> 0
   ? "Error", ::cTempDir, "can't create"
ELSE
   MakeDir( ::cTempDir+cSep+"_rels" )
   MakeDir( ::cTempDir+cSep+"docProps" )
   MakeDir( ::cTempDir+cSep+"xl" )
   MakeDir( ::cTempDir+cSep+"xl"+cSep+"_rels" )
   MakeDir( ::cTempDir+cSep+"xl"+cSep+"worksheets" )
   IF LEN(::aWorkbookDrawings)>0
		MakeDir( ::cTempDir+cSep+"xl"+cSep+"drawings" )
		MakeDir( ::cTempDir+cSep+"xl"+cSep+"drawings"+cSep+"_rels" )
		MakeDir( ::cTempDir+cSep+"xl"+cSep+"media" )
   ENDIF
// -------------------------------------------------------- [Content_Types].xml ------------------------------------------------------------------------------   
   handle := FCreate( ::cTempDir+cSep+"[Content_Types].xml" )
   FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )
   FWrite( handle, '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+chr(10) )
   FWrite( handle, '<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+chr(10) )
   FWrite( handle, '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'+chr(10) )
   FWrite( handle, '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'+chr(10) )   
   FWrite( handle, '<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+chr(10) )
   FWrite( handle, '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'+chr(10) )
   FWrite( handle, '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'+chr(10) )
   FOR I:=1 TO LEN( ::aWorkSheetNames )
		FWrite( handle, '<Override PartName="/xl/worksheets/sheet'+ALLTRIM(STR(i,10,0))+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />'+chr(10) )   
   NEXT I
   IF LEN( ::aSharedStrings ) > 0
		FWrite( handle, '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'+chr(10) )      
   ENDIF
	FOR I:=1 TO LEN(::aWorkbookDrawings )
		FWrite( handle, '<Override PartName="/xl/media/image'+ALLTRIM(STR(I,10,0))+::aWorkbookDrawings[I,3]+'" ContentType="image/'+::aWorkbookDrawings[I,3]+' />'+chr(10) )
		FWrite( handle, '<Override PartName="/xl/drawings/drawing'+ALLTRIM(STR(I,10,0))+'.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'+chr(10) )
		FWrite( handle, '<Override PartName="/xl/drawings/_rels/drawing'+ALLTRIM(STR(I,10,0))+'.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+chr(10) )
	NEXT I
   FWrite( handle, '</Types>' )	
   FClose( handle )   
   
// ------------------------------------------------------------ .rels ------------------------------------------------------------------------------   
   handle := FCreate( ::cTempDir+cSep+"_rels"+cSep+".rels" )
   FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
   FWrite( handle, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+chr(10) )
   FWrite( handle, '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'+chr(10) )
   FWrite( handle, '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'+chr(10) )   
   FWrite( handle, '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'+chr(10) )
   FWrite( handle, '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="xl/styles.xml"/>'+chr(10) )   
   FWrite( handle, '</Relationships>' )   
   FClose( handle )

   
// ------------------------------------------------------------ app.xml ------------------------------------------------------------------------------   
   handle := FCreate( ::cTempDir+cSep+"docProps"+cSep+"app.xml" )
   FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
   FWrite( handle, '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'+chr(10) )
   FWrite( handle, '<Application>'+::Application+'</Application>'+chr(10) )
   FWrite( handle, '<Company>'+::Company+'</Company>'+chr(10) )   
   FWrite( handle, '<Template></Template>'+chr(10) )      
   FWrite( handle, '<TotalTime>0</TotalTime>'+chr(10) )
   FWrite( handle, '</Properties>' )   
   FClose( handle )         

// ------------------------------------------------------------ core.xml ------------------------------------------------------------------------------   
   handle := FCreate( ::cTempDir+cSep+"docProps"+cSep+"core.xml" )
   FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
   FWrite( handle, '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'+chr(10) )
   FWrite( handle, '<dc:creator>Srdjan Dragojlovic, Kosovo is Serbia</dc:creator>'+chr(10) )
   FWrite( handle, '<cp:lastModifiedBy></cp:lastModifiedBy>'+chr(10) )
   FWrite( handle, '<dcterms:created xsi:type="dcterms:W3CDTF">'+cDateTime+'Z</dcterms:created>'+chr(10) )
   FWrite( handle, '<dcterms:modified xsi:type="dcterms:W3CDTF">'+cDateTime+'Z</dcterms:modified>'+chr(10) )
   FWrite( handle, '</cp:coreProperties>' )   
   FClose( handle )

// ------------------------------------------------------------ workbook.xml.rels ------------------------------------------------------------------------------   
   handle := FCreate( ::cTempDir+cSep+"xl"+cSep+"_rels"+cSep+"workbook.xml.rels" )
   FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
   FWrite( handle, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+chr(10) )
   FOR I:=1 TO LEN( ::aWorkSheetNames )
		FWrite( handle, '<Relationship Id="rId'+ALLTRIM(STR(i,10,0))+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'+ALLTRIM(STR(i,10,0))+'.xml"/>'+chr(10) )   
   NEXT I
   FWrite( handle, '<Relationship Id="rId'+ALLTRIM(STR(i,10,0))+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />'+chr(10) )
   IF LEN( ::aSharedStrings ) > 0
		FWrite( handle, '<Relationship Id="rId'+ALLTRIM(STR(i+1,10,0))+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'+chr(10) )
   ENDIF 	
   FWrite( handle, '</Relationships>' )   
   FClose( handle )         

// ------------------------------------------------------------ drawing.xml.rels ------------------------------------------------------------------------------   
	IF LEN( ::aWorkbookDrawings ) > 0
		handle := FCreate( ::cTempDir+cSep+"xl"+cSep+"drawings"+cSep+"_rels"+cSep+"drawing"+ALLTRIM(STR(i,10,0))+".xml.rels" )
		FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
		FWrite( handle, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+chr(10) )			
		FOR I := 1 TO LEN( ::aWorkbookDrawings )
			FWrite( handle, '<Relationship Id="rId'+ALLTRIM(STR(i,10,0))+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image"'+ALLTRIM(STR(i,10,0))+::aWorkbookDrawings[I,3]+' />'+chr(10) )
		NEXT I
		FWrite( handle, '</Relationships>' )   
		FClose( handle )         		
	ENDIF
	
// ------------------------------------------------------------ workbook.xml ------------------------------------------------------------------------------   
   handle := FCreate( ::cTempDir+cSep+"xl"+cSep+"workbook.xml" )
   FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
   FWrite( handle, '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+chr(10) )
   FWrite( handle, '<bookViews><workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/></bookViews>'+chr(10) )
   FWrite( handle, '<sheets>'+chr(10) )
   FOR I:=1 TO LEN( ::aWorkSheetNames )
		FWrite( handle, '<sheet name="'+::aWorkSheetNames[i]+'" sheetId="'+ALLTRIM(STR(i,10,0))+'" r:id="rId'+ALLTRIM(STR(i,10,0))+'"/>'+chr(10) )   
   NEXT I
   FWrite( handle, '</sheets>'+chr(10) )
   FWrite( handle, '<calcPr calcId="124519" fullCalcOnLoad="1"/>'+chr(10) )
   FWrite( handle, '</workbook>' )   
   FClose( handle )         

   // ------------------------------------------------------------ sharedStrings.xml ------------------------------------------------------------------------------------------------------------------------   
   IF ::_str_total>0
		handle := FCreate( ::cTempDir+cSep+"xl"+cSep+"sharedStrings.xml" )
		FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
        FWrite( handle, '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" Count="'+ALLTRIM(STR(::_str_total,10,0))+'" uniqueCount="'+ALLTRIM(STR(LEN(::aSharedStrings),10,0))+'">'+chr(10) )
		FOR I:=1 TO LEN(::aSharedStrings)
			FWrite( handle, '<si><t>'+::aSharedStrings[I]+'</t></si>'+chr(10) )
		NEXT I
		FWrite( handle, '</sst>' )   
		FClose( handle )         
   ENDIF 		

   // ------------------------------------------------------------ styles.xml ------------------------------------------------------------------------------------------------------------------------   
	handle := FCreate( ::cTempDir+cSep+"xl"+cSep+"styles.xml" )
	FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
    FWrite( handle, '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+chr(10) )
	FWrite( handle, '<numFmts Count="'+ALLTRIM(STR(LEN(::aNumFormat),10,0))+'">')
	FOR I:=1 TO LEN( ::aNumFormat )
		FWrite( handle, '<numFmt numFmtId="'+ALLTRIM(STR(I+163,10,0))+'" formatCode="'+::aNumFormat[I]+'"/>'+chr(10) )
	NEXT I
	FWrite( handle, '</numFmts>'+chr(10))
	FWrite( handle, '<fonts Count="'+ALLTRIM(STR(LEN(::aWorkbookFonts),10,0))+'">'+chr(10) )	
	FWrite( handle, '<font/>'+chr(10) )
	FOR I:=2 TO LEN( ::aWorkbookFonts )
	    J:= hb_ATokens( ::aWorkbookFonts[i], "," )	
		FWrite( handle, '<font>' )		
		IF j[3]=="1";FWrite( handle, '<b/>' );ENDIF
		IF j[4]=="1";FWrite( handle, '<i/>' );ENDIF
		IF j[5]=="1";FWrite( handle, '<u/>' );ENDIF
		IF j[6]=="1";FWrite( handle, '<strike/>' );ENDIF		
		FWrite( handle, '<name val="'+j[1]+'"/>' )
		FWrite( handle, '<sz val="'+j[2]+'"/>' )
		FWrite( handle, '<color rgb="'+j[7]+'"/>' )
		FWrite( handle, '</font>'+chr(10) )
	NEXT I
	FWrite( handle, '</fonts>'+chr(10) )		
	FWrite( handle, '<fills Count="'+ALLTRIM(STR(LEN(::aWorkbookFills),5,0))+'">'+chr(10) )		
	FWrite( handle, '<fill><patternFill patternType="none"/></fill>'+chr(10)+'<fill><patternFill patternType="gray125"/></fill>'+chr(10) )	
	FOR I:=3 TO LEN( ::aWorkbookFills )
	    J:= hb_ATokens( ::aWorkbookFills[i], "," )
		FWrite( handle, '<fill><patternFill patternType="'+aPattern[val(j[1])]+'"><fgColor rgb="'+j[2]+'"/><bgColor rgb="'+j[3]+'"/></patternFill></fill>'+chr(10) )
	NEXT I
	FWrite( handle, '</fills>'+chr(10) )		
	FWrite( handle, '<borders Count="'+ALLTRIM(STR(LEN(::aWorkbookBorders),5,0))+'">'+chr(10) )		
	FWrite( handle, '<border diagonalUp="false" diagonalDown="false"><left/><right/><top/><bottom/><diagonal/></border>'+chr(10) )		
	FOR I:=2 TO LEN( ::aWorkbookBorders )
	    J:= hb_ATokens( ::aWorkbookBorders[i], "," )
		FWrite( handle, '<border>' )
		FWrite( handle, iif(val(j[1])>0, '<left style="'+aBordersStyle[val(j[1])]+'"'+iif(j[2]!="000000", '><color rgb="'+j[2]+'"/></left>', '/>' ), '<left/>' ) )
		FWrite( handle, iif(val(j[3])>0, '<right style="'+aBordersStyle[val(j[3])]+'"'+iif(j[4]!="000000", '><color rgb="'+j[4]+'"/></right>', '/>' ), '<right/>' ) )
		FWrite( handle, iif(val(j[5])>0, '<top style="'+aBordersStyle[val(j[5])]+'"'+iif(j[6]!="000000", '><color rgb="'+j[6]+'"/></top>', '/>' ), '<top/>' ) )
		FWrite( handle, iif(val(j[7])>0, '<bottom style="'+aBordersStyle[val(j[7])]+'"'+iif(j[8]!="000000", '><color rgb="'+j[8]+'"/></bottom>', '/>' ), '<bottom/>' ) )		
		FWrite( handle, iif(val(j[9])>0, '<diagonal style="'+aBordersStyle[val(j[9])]+'"'+iif(j[10]!="000000", '><color rgb="'+j[10]+'"/></diagonal>', '/>' ), '<diagonal/>' ) )				
		FWrite( handle, '</border>'+chr(10) )
	NEXT I
	FWrite( handle, '</borders>'+chr(10) )		
	FWrite( handle, '<cellStyleXfs count="1">'+chr(10) )		
	FWrite( handle, '<xf />'+chr(10) )		
	FWrite( handle, '</cellStyleXfs>'+chr(10) )		
	FWrite( handle, '<cellXfs count="'+ALLTRIM(STR(LEN(::aWorkbookStyles)+2,10,0))+'">'+chr(10) )		
	FWrite( handle, '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'+chr(10) )		
	FWrite( handle, '<xf numFmtId="165"/>'+chr(10) )
	FOR I:=1 TO LEN( ::aWorkbookStyles )
	    J:= hb_ATokens( ::aWorkbookStyles[i], "," )
		FWrite( handle, '<xf ' )
		IF j[1]!="0";FWrite( handle, 'fontId="'+j[1]+'" ' );ENDIF
		IF j[2]!="0";FWrite( handle, 'borderId="'+j[2]+'" ' );ENDIF
		IF j[3]!="0";FWrite( handle, 'fillId="'+j[3]+'" ' );ENDIF
		IF j[6]!="0";FWrite( handle, 'numFmtId="'+j[6]+'" ' );ENDIF
		FWrite( handle, iif(j[4]=="0" .AND. j[5]=="0" .AND. j[7]=="0" .AND. j[8]=="0", '/>', '>' ) )
		c := IIF( j[7]!="0", 'textRotation="'+j[7]+'" ', '' )+IIF( j[8]!="0", 'wrapText="1"', '' )
		IF j[4]!="0"
			FWrite( handle, '<alignment horizontal="'+aHorizontal[val(j[4])]+iif( j[5]=="0", '" '+c+'/>', '" ' ) )
		ENDIF
		IF j[5]!="0"
			FWrite( handle, IIF(j[4]=="0", '<alignment ', ' '+'vertical="'+aVertical[val(j[5])]+'" '+c+'/>' ) )
		ENDIF
		IF j[4]!="0" .OR. j[5]!="0" .OR. j[7]!="0" .OR. j[8]!="0"; FWrite( handle, '</xf>'+chr(10) );ENDIF
	NEXT I	
	FWrite( handle, '</cellXfs>'+chr(10) )		
	FWrite( handle, '</styleSheet>' )   
	FClose( handle )         
   // -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	FOR I:=1 TO LEN( ::aWorkSheetNames )
		::aWorkSheetObjects[i]:WriteWorksheet(i, ::cTempDir+cSep+"xl"+cSep+"worksheets"+cSep)
	NEXT I
	c := DiskName() + hb_OSDriveSeparator() + hb_PS() + CurDir()
	DirChange( ::cTempDir )
	FERASE( c+cSep+::cName )
	myzip( c+cSep+::cName, "*.*" )
	DirChange( c )
	//hb_DirRemoveAll( ::cTempDir )
ENDIF
Return Self

CLASS WorkSheet
DATA cName
DATA oParent
DATA paperSize INIT 9
DATA lLandscape INIT .F.
DATA horizontalDpi INIT 300
DATA verticalDpi INIT 300
DATA leftMargin INIT 0.2
DATA rightMargin INIT 0.2
DATA topMargin INIT 0.2
DATA bottomMargin INIT 0.2
DATA headerMargin INIT 0.1
DATA footerMargin INIT 0.1
DATA fitToPage INIT .F.
DATA fitToHeight INIT .F.
DATA fitToWidth INIT .F.
DATA zoom INIT 100
DATA zoom_scale_normal INIT .T.
DATA print_scale INIT 100
DATA right_to_left INIT .F.
DATA show_zeros INIT .T.
DATA leading_zeros INIT .F.
DATA nMaxRow PROTECTED
DATA nMaxCol PROTECTED
DATA aData PROTECTED
DATA lChar PROTECTED
DATA aMergeCells PROTECTED
DATA aStyle PROTECTED
DATA aCols PROTECTED
DATA aRows PROTECTED
DATA aCharts PROTECTED
DATA lHeader PROTECTED
DATA lFooter PROTECTED
DATA aHeader PROTECTED
DATA aFooter PROTECTED
METHOD New(cName)
METHOD WriteWorksheet(n, cPath)
METHOD Cell(uAddr, nValue, nStyle)
METHOD MergeCell(uAddr)
METHOD ColumnsWidth( nMin, nLast, nWidth )
METHOD AddHeader( cLeft, cCenter, cRight )
METHOD AddFooter( cLeft, cCenter, cRight )
METHOD RowDetail( nRow, nHeight, nStyle, lHide )
METHOD AddChart( uAddr, oChart, nX, nY, nX_Scale, nY_Scale )
METHOD AddDrawing( uAddr, cDrawingPath, nX, nY, nX_Scale, nY_Scale )
ENDCLASS

METHOD New(cName) CLASS WorkSheet
::cName := cName
::aData := {}
::aMergeCells := {}
::aCols := {}
::aRows := {}
::aStyle := {}
::aCharts := {}
::nMaxRow := 0
::nMaxCol := 0
::lChar := .F.
::lHeader := .F.
::lFooter := .F.
::aHeader := { "", "", "" }
::aFooter := { "", "", "" }
Return Self

METHOD RowDetail( nRow, nHeight, nStyle, lHide ) CLASS WorkSheet
RETURN Self

METHOD AddChart( uAddr, oChart, nX, nY, nX_Scale, nY_Scale ) CLASS WorkSheet
RETURN Self

METHOD AddDrawing( uAddr, cDrawingPath, nX, nY, nX_Scale, nY_Scale ) CLASS WorkSheet
RETURN Self

METHOD MergeCell( uAddr ) CLASS WorkSheet
LOCAL nCol1, nCol2, nRow1, nRow2, adr, I, J, K
IF HB_ISSTRING(uAddr) .AND. ASCAN( ::aMergeCells, uAddr )==0
	AADD(  ::aMergeCells, uAddr )
	adr := hb_ATokens( uAddr, ":" )
	WORKSHEET_RC( adr[1], @nRow1, @nCol1 )
	WORKSHEET_RC( adr[2], @nRow2, @nCol2 )
    FOR I:=nRow1 TO nRow2
		FOR J:=nCol1 TO nCol2
			::nMaxCol := IIF( J > ::nMaxCol, J, ::nMaxCol )
			::nMaxRow := IIF( I > ::nMaxRow, I, ::nMaxRow )
			IF LEN( ::aData ) < I
				FOR k := LEN(::aData)+1 TO I 
					AADD( ::aData, {} )
				NEXT k
			ENDIF
			IF LEN( ::aData[I] ) < J
				FOR k := LEN(::aData[I])+1 TO J
					AADD( ::aData[I], NIL )
				NEXT k	   
			ENDIF
			IF ::aData[I,J]==NIL
				::aData[I,J] := CHR(0)
			ENDIF
		NEXT J
	NEXT I
ENDIF
Return Self

METHOD AddHeader( cLeft, cCenter, cRight )
IF HB_ISSTRING( cLeft )
   ::aHeader[1] := ReplaceAmp(cLeft)
ENDIF
IF HB_ISSTRING( cCenter )
   ::aHeader[2] := ReplaceAmp(cCenter)
ENDIF
IF HB_ISSTRING( cRight )
   ::aHeader[3] := ReplaceAmp(cRight)
ENDIF
::lHeader := .T.
RETURN Self

METHOD AddFooter( cLeft, cCenter, cRight )
IF HB_ISSTRING( cLeft )
   ::aFooter[1] := ReplaceAmp(cLeft)
ENDIF
IF HB_ISSTRING( cCenter )
   ::aFooter[2] := ReplaceAmp(cCenter)
ENDIF
IF HB_ISSTRING( cRight )
   ::aFooter[3] := ReplaceAmp(cRight)
ENDIF
::lFooter := .T.
RETURN Self

METHOD ColumnsWidth( nMin, nMax, nWidth ) CLASS WorkSheet
IF HB_ISNUMERIC(nMin) .AND. HB_ISNUMERIC(nMax) .AND. HB_ISNUMERIC(nWidth)
	AADD( ::aCols, { nMin, nMax, nWidth } )
ENDIF
Return Self

METHOD Cell( uAddr, xValue, nStyle ) CLASS WorkSheet
LOCAL nCol := 0, nRow := 0, i, j, k, l, lIsMerge := .F., nRow1 := 0, nCol1 := 0, nRow2 := 0, nCol2 := 0

IF HB_ISARRAY( uAddr )
	nRow := uAddr[1]
	nCol := uAddr[2]
ELSE
	WORKSHEET_RC( uAddr, @nRow, @nCol )
ENDIF
::nMaxCol := IIF( nCol > ::nMaxCol, nCol, ::nMaxCol )
::nMaxRow := IIF( nRow > ::nMaxRow, nRow, ::nMaxRow )
IF LEN( ::aData ) < nRow
	FOR i := LEN(::aData)+1 TO nRow 
		AADD( ::aData, {} )
	NEXT I
ENDIF
IF LEN( ::aData[nRow] ) < nCol
	FOR i := LEN(::aData[nRow])+1 TO nCol
		AADD( ::aData[nRow], NIL )
	NEXT I	   
ENDIF
lIsMerge := IIF( HB_ISSTRING(::aData[nRow,nCol]) .AND. ::aData[nRow,nCol]==CHR(0),.T.,.F.)
IF HB_ISSTRING(xValue)
	nPos := AT( "&", xValue )
	xValue := IIF( nPos>0, LEFT(xValue,nPos)+"amp;"+SUBSTR(xValue,nPos+1), xValue )
ENDIF
IF !HB_ISNIL(xValue)
	::aData[nRow,nCol] := xValue
	IF HB_ISSTRING(xValue) .AND. LEFT(xValue,1) != "="
	    ::oParent:_str_total++
		i := ASCAN( ::oParent:aSharedStrings, xValue )
		IF i==0
			AADD( ::oParent:aSharedStrings, xValue )
		ENDIF
	ENDIF
ENDIF
IF !HB_ISNIL(nStyle) .AND. HB_ISNUMERIC(nStyle) 	
	IF LEN( ::aStyle ) < nRow
		FOR i := LEN(::aStyle)+1 TO nRow 
			AADD( ::aStyle, {} )
		NEXT I
	ENDIF
	IF LEN( ::aStyle[nRow] ) < nCol
		FOR i := LEN(::aStyle[nRow])+1 TO nCol
			AADD( ::aStyle[nRow], NIL )
		NEXT I	   
	ENDIF
	::aStyle[nRow,nCol] := ALLTRIM(STR(nStyle,10,0))
	IF lIsMerge
		FOR l:=1 TO LEN( ::aMergeCells )
			adr := hb_ATokens( ::aMergeCells[l], ":" )
			WORKSHEET_RC( adr[1], @nRow1, @nCol1 )
		    IF nRow1==nRow .AND. nCol1==nCol
				WORKSHEET_RC( adr[2], @nRow2, @nCol2 )
				FOR I:=nRow1 TO nRow2
					FOR J:=nCol1 TO nCol2
						IF LEN( ::aStyle ) < I
							FOR k := LEN(::aStyle)+1 TO I 
								AADD( ::aStyle, {} )
							NEXT k
						ENDIF
						IF LEN( ::aStyle[I] ) < J
							FOR k := LEN(::aStyle[I])+1 TO J
								AADD( ::aStyle[I], NIL )
							NEXT k	   
						ENDIF
						::aStyle[I,J] := ALLTRIM(STR(nStyle,10,0))						
					NEXT J
				NEXT I
				EXIT
			ENDIF
		NEXT l
	ENDIF
ENDIF
Return ::aData[nRow,nCol]

METHOD WriteWorksheet(n, cPath) CLASS WorkSheet
LOCAL i, j, c, x
handle := FCreate( cPath+"sheet"+alltrim(str(n,10,0))+".xml" )
FWrite( handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+chr(10) )   
FWrite( handle, '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+chr(10) )
FWrite( handle, '<sheetPr><pageSetUpPr fitToPage="'+iif(::fitToPage,"1","0")+'"/></sheetPr>'+chr(10) )
IF LEN( ::aCols ) > 0
	FWrite( handle, '<cols>'+chr(10) )
	FOR I=1 TO LEN( ::aCols )
		FWrite( handle, '<col min="'+ALLTRIM(STR(::aCols[I,1],5,0))+'" max="'+ALLTRIM(STR(::aCols[I,2],5,0))+'" width="'+ALLTRIM(STR(::aCols[I,3],18,8))+'" bestFit="1" customWidth="1"/>'+chr(10) ) 
	NEXT I
	FWrite( handle, '</cols>'+chr(10) )
ENDIF
FWrite( handle, '<sheetData>'+chr(10) )
FOR I:=1 TO LEN( ::aData )
	IF LEN(::aData[I])>0 
	    FOR J:=1 TO LEN(::aData[I])
		    IF !HB_ISNIL(::aData[I,J])
				c := J
				J := LEN(::aData[I]) + 1
            ENDIF
        NEXT J 
		FWrite( handle, '<row r="'+ALLTRIM(STR(I,10,0))+'" spans="'+ALLTRIM(STR(c,10,0))+':'+ALLTRIM(STR(LEN(::aData[I]),10,0))+'">'+chr(10) )
		FOR J:=1 TO LEN(::aData[I])
			IF !HB_ISNIL(::aData[I,J])
				c := ColumnIndexToColumnLetter(J)+ALLTRIM(STR(I,10,0))
				IF HB_ISNUMERIC(::aData[I,J])			
                    v := ALLTRIM( BestPrecision(::aData[I,J]) )
					IF !(LEN( ::aStyle ) < I) .AND. !(LEN(::aStyle[I])<J) .AND. (HB_ISSTRING(::aStyle[I,J]))
						FWrite( handle, '<c r="'+c+'" s="'+::aStyle[I,J]+'"><v>'+v+'</v></c>'+chr(10) )
					ELSE
						FWrite( handle, '<c r="'+c+'"><v>'+v+'</v></c>'+chr(10) )
					ENDIF
				ELSEIF HB_ISSTRING(::aData[I,J])
					IF LEFT( ::aData[I,J], 1 ) == "="
						v := SUBSTR(::aData[I,J],2)
						IF !(LEN( ::aStyle ) < I) .AND. !(LEN(::aStyle[I])<J) .AND. (HB_ISSTRING(::aStyle[I,J]))
							FWrite( handle, '<c r="'+c+'" t="str" s="'+::aStyle[I,J]+'"><f>'+v+'</f></c>'+chr(10) )
						ELSE
							FWrite( handle, '<c r="'+c+'" t="str"><f>'+v+'</f></c>'+chr(10) )
						ENDIF
					ELSEIF LEFT( ::aData[I,J], 1 ) == CHR(0)
						IF !(LEN( ::aStyle ) < I) .AND. !(LEN(::aStyle[I])<J) .AND. (HB_ISSTRING(::aStyle[I,J]))
							FWrite( handle, '<c r="'+c+'" s="'+::aStyle[I,J]+'"></c>'+chr(10) )
						ENDIF							
					ELSE
						v := ALLTRIM( STR( ASCAN( ::oParent:aSharedStrings, ::aData[I,J] )-1, 10, 0 ) )
						IF !(LEN( ::aStyle ) < I) .AND. !(LEN(::aStyle[I])<J) .AND. (HB_ISSTRING(::aStyle[I,J]))
							FWrite( handle, '<c r="'+c+'" t="s" s="'+::aStyle[I,J]+'"><v>'+v+'</v></c>'+chr(10) )
						ELSE
							FWrite( handle, '<c r="'+c+'" t="s"><v>'+v+'</v></c>'+chr(10) )
						ENDIF
					ENDIF
				ELSEIF HB_ISLOGICAL(::aData[I,J])						
                    v := IIF( ::aData[I,J], "1", "0" )
					IF !(LEN( ::aStyle ) < I) .AND. !(LEN(::aStyle[I])<J) .AND. (HB_ISSTRING(::aStyle[I,J]))
						FWrite( handle, '<c r="' +c+ '" t="b" s="'+::aStyle[I,J]+'"><v>' +v+ '</v></c>'+chr(10) )
					ELSE
						FWrite( handle, '<c r="' +c+ '" t="b"><v>' +v+ '</v></c>'+chr(10) )							
					ENDIF
				ELSEIF HB_ISDATE(::aData[I,J])	
					x := Set( _SET_DATEFORMAT, "dd.mm.yyyy" ) 
                    v := ALLTRIM( STR(::aData[I,J]-CTOD( "01.01.1900" )+2,10,0) )
					FWrite( handle, '<c r="' +c+ '" s="1"><v>' +v+ '</v></c>'+chr(10) )
					Set( _SET_DATEFORMAT, x )
				ENDIF
			ENDIF
		NEXT J
		FWrite( handle, '</row>'+chr(10) )
	ENDIF
NEXT I
FWrite( handle, '</sheetData>'+chr(10) )
IF LEN( ::aMergeCells ) > 0
	FWrite( handle, '<mergeCells>'+chr(10) )
	FOR I:=1 TO LEN( ::aMergeCells )
		FWrite( handle, '<mergeCell ref="'+::aMergeCells[I]+'"/>'+chr(10) )
	NEXT I
	FWrite( handle, '</mergeCells>'+chr(10) )
ENDIF
FWrite( handle, '<pageMargins left="'+alltrim(str(::leftMargin,10,2))+'" right="'+alltrim(str(::rightMargin,10,2))+'" top="'+alltrim(str(::topMargin,10,2))+'" bottom="'+alltrim(str(::bottomMargin,10,2))+'" header="'+alltrim(str(::headerMargin,10,2))+'" footer="'+alltrim(str(::footerMargin,10,2))+'"/>'+chr(10) )
FWrite( handle, '<pageSetup paperSize="'+alltrim(str(::PaperSize,10,0))+'" horizontalDpi="'+alltrim(str(::horizontalDpi,10,2))+'" verticalDpi="'+alltrim(str(::verticalDpi,10,2))+'"'+iif(::lLandscape,' orientation="landscape"',' orientation="portrait"')+'/>'+chr(10) )
IF ::lHeader .OR. ::lFooter
	FWrite( handle, '<headerFooter>' )
	IF ::lHeader
		FWrite( handle, '<oddHeader>' )
		IF LEN(::aHeader[1])>0
			FWrite( handle, '&amp;L'+::aHeader[1] )
		ENDIF
		IF LEN(::aHeader[2])>0
			FWrite( handle, '&amp;C'+::aHeader[2] )
		ENDIF
		IF LEN(::aHeader[3])>0
			FWrite( handle, '&amp;R'+::aHeader[3] )
		ENDIF		
		FWrite( handle, '</oddHeader>' )
	ENDIF
	IF ::lFooter
		FWrite( handle, '<oddFooter>' )
		IF LEN(::aFooter[1])>0
			FWrite( handle, '&amp;L'+::aFooter[1] )
		ENDIF
		IF LEN(::aFooter[2])>0
			FWrite( handle, '&amp;C'+::aFooter[2] )
		ENDIF
		IF LEN(::aFooter[3])>0
			FWrite( handle, '&amp;R'+::aFooter[3] )
		ENDIF		
		FWrite( handle, '</oddFooter>' )
	ENDIF	
	FWrite( handle, '</headerFooter>' )
ENDIF
FWrite( handle, '</worksheet>' )   
FClose( handle )         
Return Self

STATIC Function ColumnIndexToColumnLetter(colIndex)
LOCAL div := colIndex
LOCAL colLetter := ""
LOCAL modnum := 0
While div > 0
    modnum := MOD((div - 1),26)
    colLetter := Chr(65 + modnum) + colLetter
    div := Int((div - modnum) / 26)
End 
Return colLetter

STATIC FUNCTION BestPrecision(n)
LOCAL I, p:=0
IF INT(n)<>n
	FOR I:=1 TO 17
		IF n<(10^I)
			p := 17-i
			EXIT
		ENDIF
	NEXT I
ENDIF
RETURN STR( n, 18, p )

STATIC PROCEDURE MyZip( cName, cWild )

   LOCAL hZip, aDir, aFile, cZipName, cPath, cFileName, cExt, lUnicode := .T.
   hb_FNameSplit( cName, @cPath, @cFileName, @cExt )
   cZipName := hb_FNameMerge( cPath, cFileName, cExt )
   
	hZip := hb_zipOpen( cZipName )
	IF ! Empty( hZip )
        IF ! Empty( cWild )
            hb_FNameSplit( cWild, @cPath, @cFileName, @cExt )
            aDir := hb_DirScan( cPath, cFileName + cExt )
            FOR EACH aFile IN aDir
               IF ! cPath + aFile[ 1 ] == cZipName
                  hb_zipStoreFile( hZip, cPath + aFile[ 1 ], cPath + aFile[ 1 ],,, lUnicode )
               ENDIF
            NEXT
        ENDIF
		hb_zipClose( hZip )
	ENDIF

RETURN

STATIC FUNCTION ReplaceAmp(xValue)
LOCAL nPos, c:= ""
WHILE (nPos := AT( "&", xValue ))>0
	c += LEFT(xValue,nPos)+"amp;"
	xValue := SUBSTR(xValue,nPos+1)	
END
c+=xValue
RETURN c

#pragma BEGINDUMP
#include "hbapi.h"

HB_FUNC( WORKSHEET_RC )
{ 
  char *cellAddr = hb_parc(1);
  int ii=0, jj, colVal=0;
  while(cellAddr[ii++] >= 'A') {};
  ii--;
  for(jj=0;jj<ii;jj++) colVal = 26*colVal + toupper(cellAddr[jj]) -'A' + 1;
  hb_storni( atoi(cellAddr+ii), 2 );
  hb_storni( colVal, 3 );
}	

#pragma ENDDUMP
