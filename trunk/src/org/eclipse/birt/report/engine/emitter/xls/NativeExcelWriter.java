package org.eclipse.birt.report.engine.emitter.xls;

import java.awt.Color;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.eclipse.birt.report.engine.content.IHyperlinkAction;
import org.eclipse.birt.report.engine.content.IReportContent;
import org.eclipse.birt.report.engine.emitter.xls.layout.ExcelContext;

public class NativeExcelWriter implements IExcelWriter
{

    public static final int rightToLeftisTrue = 1; // bidi_acgc added
    private final POIWriterXLS writer = new POIWriterXLS();

    public POIWriterXLS getWriter()
    {
        return writer;
    }

    private String pageHeader, pageFooter;
    private int sheetIndex = 1;

    protected static Logger logger = Logger.getLogger( NativeExcelWriter.class.getName() );

    ExcelContext context = null;

    public NativeExcelWriter( ExcelContext context )
    {
        this( "UTF-8", context );
    }

    public NativeExcelWriter( OutputStream out )
    {
        writer.open( out, "UTF-8" );
    }

    public NativeExcelWriter( OutputStream out, ExcelContext context )
    {
        this.context = context;
        writer.open( out, "UTF-8", context );
    }

    public NativeExcelWriter( String encoding, ExcelContext context )
    {
        this( context.getOutputSteam(), context );
    }

    private void writeDocumentProperties( IReportContent reportContent )
    {
        if ( reportContent == null ) { return; }
        /*	ReportDesignHandle reportDesign = reportContent.getDesign( )
        	.getReportDesign( );
        	writer.openTag( "DocumentProperties" );
        	writer.attribute( "xmlns", "urn:schemas-microsoft-com:office:office" );
        	writer.openTag( "Author" );
        	writer
        	.text( reportDesign
        			.getStringProperty( IModuleModel.AUTHOR_PROP ) );
        	writer.closeTag( "Author" );
        	writer.openTag( "Title" );
        	writer.text( reportContent.getTitle( ) );
        	writer.closeTag( "Title" );
        	writer.openTag( "Description" );
        	writer.text( reportDesign.getComments( ) );
        	writer.closeTag( "Description" );
        	writer.openTag( "Subject" );
        	writer.text( reportDesign.getSubject( ) );
        	writer.closeTag( "Subject" );
        	writer.closeTag( "DocumentProperties" );
        	*/

    }

    private void writeText( int type, Object value, StyleEntry style )
    {
        writer.writeText( type, value, style );
    }

    public void startRow( double rowHeight )
    {
        writer.createRow( (float)rowHeight );
    }

    public void endRow()
    {
    }

    public void outputData( String sheet, SheetData sheetData, StyleEntry style, int column, int colSpan )
    {
        // TODO: ignore sheet here. If this function is needed, need to
        // implement.
        outputData( sheetData, style, column, colSpan );
    }

    public void outputData( SheetData sheetData, StyleEntry style, int column, int colSpan )
    {
        int rowSpan = sheetData.getRowSpan();
        int styleId = sheetData.getStyleId();
        int type = sheetData.getDataType();
        if ( type == SheetData.IMAGE )
        {
            if ( sheetData instanceof ImageData )
            {
                ImageData imageData = (ImageData)sheetData;
                outputData( Data.STRING, imageData, style, column, colSpan, sheetData.getRowSpan(), sheetData.getStyleId(), null, null );
            }
        }
        else
        {
            Data d = (Data)sheetData;
            Object value = d.getValue();
            HyperlinkDef hyperLink = d.getHyperlinkDef();
            BookmarkDef linkedBookmark = d.getLinkedBookmark();
            outputData( type, value, style, column, colSpan, rowSpan, styleId, hyperLink, linkedBookmark );
        }
    }

    public void outputData( int col, int row, int type, Object value )
    {
        outputData( type, value, null, col, 0, 0, -1, null, null );
    }

    private void outputData( int type, Object value, StyleEntry style, int column, int colSpan, int rowSpan, int styleId, HyperlinkDef hyperLink, BookmarkDef linkedBookmark )
    {

        String urlAddress = getURLAddress( hyperLink, linkedBookmark );

        writer.createCell( column, colSpan, rowSpan, styleId, style, hyperLink, urlAddress );
        
        if ( value != null & value instanceof ImageData )
        {
            writeImage( type, (ImageData)value, style, column, colSpan, rowSpan );
        }
        else
        {
            writeText( type, value, style );
        }
        
    }

    private void writeImage( int type, ImageData imageData, StyleEntry style, int column, int colSpan, int rowSpan )
    {
        writer.writeImage( type, imageData, style, column, colSpan, rowSpan );
    }

    protected String getURLAddress( HyperlinkDef hyperLink, BookmarkDef linkedBookmark )
    {
        String urlAddress = null;
        if ( hyperLink != null )
        {
            urlAddress = hyperLink.getUrl();
            if ( hyperLink.getType() == IHyperlinkAction.ACTION_BOOKMARK )
            {
                if ( linkedBookmark != null )
                    urlAddress = "#" + linkedBookmark.getValidName();
                else
                {
                    logger.log( Level.WARNING, "The bookmark: {" + urlAddress + "} is not defined!" );
                }
            }
            if ( urlAddress != null && urlAddress.length() >= 255 )
            {
                logger.log( Level.WARNING, "The URL: {" + urlAddress + "} is too long!" );
                urlAddress = urlAddress.substring( 0, 254 );
            }

        }
        return urlAddress;
    }

    protected void writeComments( HyperlinkDef linkDef )
    {
        /*
        String toolTip = linkDef.getToolTip( );
        writer.openTag( "Comment" );
        writer.openTag( "ss:Data" );
        writer.attribute( "xmlns", "http://www.w3.org/TR/REC-html40" );
        writer.openTag( "Font" );
        //		writer.attribute( "html:Face", "Tahoma" );
        //		writer.attribute( "x:CharSet", "1" );
        //		writer.attribute( "html:Size", "8" );
        //		writer.attribute( "html:Color", "#000000" );
        writer.text( toolTip );
        writer.closeTag( "Font" );
        writer.closeTag( "ss:Data" );
        writer.closeTag( "Comment" );
        */
    }

    

    private void writeAlignment( String horizontal, String vertical, float indent, String direction, boolean wrapText )
    {
        /*
        writer.openTag( "Alignment" );

        if ( isValid( horizontal ) )
        {
        	writer.attribute( "ss:Horizontal", horizontal );
        }

        if ( isValid( vertical ) )
        {
        	writer.attribute( "ss:Vertical", vertical );
        }
        if ( indent != 0f )
        {
        	writer.attribute( "ss:Indent", indent );	
        }
        if ( isValid( direction ) )
        {
        	if ( CSSConstants.CSS_RTL_VALUE.equals( direction ) )
        		writer.attribute( "ss:ReadingOrder", "RightToLeft" );
        	else
        		writer.attribute( "ss:ReadingOrder", "LeftToRight" );
        }
        if(wrapText)
        {
        	writer.attribute( "ss:WrapText", "1" );
        }

        writer.closeTag( "Alignment" );
        */
    }

    private void writeBorder( String position, String lineStyle, Integer weight, Color color )
    {
        /*
        writer.openTag( "Border" );
        writer.attribute( "ss:Position", position );
        if ( isValid( lineStyle ) )
        {
        	writer.attribute( "ss:LineStyle", lineStyle );
        }

        if ( weight != null && weight > 0 )
        {
        	writer.attribute( "ss:Weight", weight );
        }

        if ( color != null )
        {
        	writer.attribute( "ss:Color", toString( color ) );
        }

        writer.closeTag( "Border" );
        */
    }

    private void writeFont( String fontName, Float size, Boolean bold, Boolean italic, Boolean strikeThrough, Boolean underline, Color color )
    {
        /*
        writer.openTag( "Font" );

        if ( isValid( fontName ) )
        {
        	fontName = getFirstFont( fontName );
        	writer.attribute( "ss:FontName", fontName );
        }

        if ( size != null )
        {
        	writer.attribute( "ss:Size", size );
        }

        if ( bold != null && bold )
        {
        	writer.attribute( "ss:Bold", 1 );
        }

        if ( italic != null && italic )
        {
        	writer.attribute( "ss:Italic", 1 );
        }

        if ( strikeThrough != null && strikeThrough )
        {
        	writer.attribute( "ss:StrikeThrough", 1 );
        }

        if ( underline != null && underline )
        {
        	writer.attribute( "ss:Underline", "Single" );
        }

        if ( color != null )
        {
        	writer.attribute( "ss:Color", toString( color ) );
        }

        writer.closeTag( "Font" );
        */
    }

    private void writeBackGroudColor( StyleEntry style )
    {
        /*
        Color bgColor = (Color) style
        		.getProperty( StyleConstant.BACKGROUND_COLOR_PROP );
        if ( bgColor != null )
        {
        	writer.openTag( "Interior" );
        	writer.attribute( "ss:Color", toString( bgColor ) );
        	writer.attribute( "ss:Pattern", "Solid" );
        	writer.closeTag( "Interior" );
        }
        */
    }

    private boolean isValid( String value )
    {
        return !StyleEntry.isNull( value );
    }

    private String getFirstFont( String fontName )
    {
        int firstSeperatorIndex = fontName.indexOf( ',' );
        if ( firstSeperatorIndex != -1 )
        {
            return fontName.substring( 0, firstSeperatorIndex );
        }
        else
        {
            return fontName;
        }
    }

    private void declareStyle( StyleEntry style, int id )
    {
        /*
        boolean wrapText = context.getWrappingText( );
        String whiteSpace = (String) style
        		.getProperty( StyleConstant.WHITE_SPACE );
        if ( CSSConstants.CSS_NOWRAP_VALUE.equals( whiteSpace ) )
        {
        	wrapText = false;
        }

        writer.openTag( "Style" );
        writer.attribute( "ss:ID", id );
        if ( style.isHyperlink( ) )
        {
        	writer.attribute( "ss:Parent", "HyperlinkId" );
        }

        if ( id >= StyleEngine.RESERVE_STYLE_ID )
        {
        	String direction = (String) style
        			.getProperty( StyleConstant.DIRECTION_PROP ); // bidi_hcg
        	String horizontalAlign = (String) style
        			.getProperty( StyleConstant.H_ALIGN_PROP );
        	String verticalAlign = (String) style
        			.getProperty( StyleConstant.V_ALIGN_PROP );
        	float indent = ExcelUtil.convertTextIndentToEM(
        			(FloatValue) style.getProperty( StyleConstant.TEXT_INDENT ),
        			(Float) style.getProperty( StyleConstant.FONT_SIZE_PROP ) );
        	writeAlignment( horizontalAlign, verticalAlign, indent, direction, wrapText );
        	writer.openTag( "Borders" );
        	Color bottomColor = (Color) style
        			.getProperty( StyleConstant.BORDER_BOTTOM_COLOR_PROP );
        	String bottomLineStyle = (String) style
        			.getProperty( StyleConstant.BORDER_BOTTOM_STYLE_PROP );
        	Integer bottomWeight = (Integer) style
        			.getProperty( StyleConstant.BORDER_BOTTOM_WIDTH_PROP );
        	writeBorder( "Bottom", bottomLineStyle, bottomWeight, bottomColor );

        	Color topColor = (Color) style
        			.getProperty( StyleConstant.BORDER_TOP_COLOR_PROP );
        	String topLineStyle = (String) style
        			.getProperty( StyleConstant.BORDER_TOP_STYLE_PROP );
        	Integer topWeight = (Integer) style
        			.getProperty( StyleConstant.BORDER_TOP_WIDTH_PROP );
        	writeBorder( "Top", topLineStyle, topWeight, topColor );

        	Color leftColor = (Color) style
        			.getProperty( StyleConstant.BORDER_LEFT_COLOR_PROP );
        	String leftLineStyle = (String) style
        			.getProperty( StyleConstant.BORDER_LEFT_STYLE_PROP );
        	Integer leftWeight = (Integer) style
        			.getProperty( StyleConstant.BORDER_LEFT_WIDTH_PROP );
        	writeBorder( "Left", leftLineStyle, leftWeight, leftColor );

        	Color rightColor = (Color) style
        			.getProperty( StyleConstant.BORDER_RIGHT_COLOR_PROP );
        	String rightLineStyle = (String) style
        			.getProperty( StyleConstant.BORDER_RIGHT_STYLE_PROP );
        	Integer rightWeight = (Integer) style
        			.getProperty( StyleConstant.BORDER_RIGHT_WIDTH_PROP );
        	writeBorder( "Right", rightLineStyle, rightWeight, rightColor );

        	Color diagonalColor = (Color) style
        			.getProperty( StyleConstant.BORDER_DIAGONAL_COLOR_PROP );
        	String diagonalStyle = (String) style
        			.getProperty( StyleConstant.BORDER_DIAGONAL_STYLE_PROP );
        	Integer diagonalWidth = (Integer) style
        			.getProperty( StyleConstant.BORDER_DIAGONAL_WIDTH_PROP );
        	writeBorder( "DiagonalLeft", diagonalStyle, diagonalWidth,
        			diagonalColor );

        	writer.closeTag( "Borders" );

        	String fontName = (String) style
        			.getProperty( StyleConstant.FONT_FAMILY_PROP );
        	Float size = (Float) style
        			.getProperty( StyleConstant.FONT_SIZE_PROP );
        	Boolean fontStyle = (Boolean) style
        			.getProperty( StyleConstant.FONT_STYLE_PROP );
        	Boolean fontWeight = (Boolean) style
        			.getProperty( StyleConstant.FONT_WEIGHT_PROP );
        	Boolean strikeThrough = (Boolean) style
        			.getProperty( StyleConstant.TEXT_LINE_THROUGH_PROP );
        	Boolean underline = (Boolean) style
        			.getProperty( StyleConstant.TEXT_UNDERLINE_PROP );
        	Color color = (Color) style.getProperty( StyleConstant.COLOR_PROP );
        	writeFont( fontName, size, fontWeight, fontStyle, strikeThrough,
        			underline, color );
        	writeBackGroudColor( style );
        }

        writeDataFormat( style );

        writer.closeTag( "Style" );
        */
    }

    private String toString( Color color )
    {
        if ( color == null )
            return null;
        return "#" + toHexString( color.getRed() ) + toHexString( color.getGreen() ) + toHexString( color.getBlue() );
    }

    private static String toHexString( int c )
    {
        String result = Integer.toHexString( c );
        if ( result.length() < 2 )
        {
            result = "0" + result;
        }
        return result;
    }

    // here the user input can be divided into two cases :
    // the case in the birt input like G and the Currency
    // the case in excel format : like 0.00E00

    private void writeDeclarations()
    {
        /*
        writer.startWriter( );
        writer.println( );
        writer.println( "<?mso-application progid=\"Excel.Sheet\"?>" );

        writer.openTag( "Workbook" );

        writer.attribute( "xmlns",
        "urn:schemas-microsoft-com:office:spreadsheet" );
        writer.attribute( "xmlns:o", "urn:schemas-microsoft-com:office:office" );
        writer.attribute( "xmlns:x", "urn:schemas-microsoft-com:office:excel" );
        writer.attribute( "xmlns:ss",
        "urn:schemas-microsoft-com:office:spreadsheet" );
        writer.attribute( "xmlns:html", "http://www.w3.org/TR/REC-html40" );
        */
    }

    public void startSheet( String name )
    {
        startSheet( name, null );
    }

    public void startSheet( String name, double[] coordinates )
    {
        writer.startSheet( name, coordinates, context.isRTL() );

    }

    public void closeSheet()
    {

    }

    public void outputColumns( double[] width )
    {

    }

    public void endTable()
    {

    }

    public void insertHorizontalMargin( int height, int span )
    {
        /*
        writer.openTag( "Row" );
        writer.attribute( "ss:AutoFitHeight", 0 );
        writer.attribute( "ss:Height", height );

        writer.openTag( "Cell" );
        writer.attribute( " ss:MergeAcross", span );
        writer.closeTag( "Cell" );

        writer.closeTag( "Row" );
        */
    }

    public void insertVerticalMargin( int start, int end, int length )
    {
        /*
        writer.openTag( "Row" );
        writer.attribute( "ss:AutoFitHeight", 0 );
        writer.attribute( "ss:Height", 1 );

        writer.openTag( "Cell" );
        writer.attribute( "ss:Index", start );
        writer.attribute( " ss:MergeDown", length );
        writer.closeTag( "Cell" );

        writer.openTag( "Cell" );
        writer.attribute( "ss:Index", end );
        writer.attribute( " ss:MergeDown", length );
        writer.closeTag( "Cell" );

        writer.closeTag( "Row" );
        */
    }

    public void startSheet( double[] coordinates, String pageHeader, String pageFooter, String name )
    {
        this.pageHeader = pageHeader;
        this.pageFooter = pageFooter;
        startSheet( name, coordinates );
        sheetIndex += 1;
    }

    public void endSheet( double[] coordinates, String orientation, int pageWidth, int pageHeight, float leftMargin, float rightMargin, float topMargin, float bottomMargin )
    {
        endTable();
        writer.declareWorkSheetOptions( orientation, pageWidth, pageHeight, leftMargin, rightMargin, topMargin, bottomMargin, pageHeader, pageFooter );
        closeSheet();
    }

    public void start( IReportContent report, Map< StyleEntry, Integer > styles,
    // TODO: style ranges.
    // List<ExcelRange> styleRanges,
            HashMap< String, BookmarkDef > bookmarkList )
    {
        writeDeclarations();
        writeDocumentProperties( report );

    }

    public void end()
    {

    }

    public void close()
    {
        writer.close();

    }

    public void setSheetIndex( int sheetIndex )
    {
        this.sheetIndex = sheetIndex;
    }

    public void endSheet()
    {
        endSheet( null, null, 0, 0, 0, 0, 0, 0 );
    }

    public void startRow()
    {
        startRow( -1 );
    }

    public String defineName( String cells )
    {
        return null;
    }
}
