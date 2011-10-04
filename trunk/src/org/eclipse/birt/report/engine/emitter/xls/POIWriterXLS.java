package org.eclipse.birt.report.engine.emitter.xls;

import java.awt.Color;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;
import org.eclipse.birt.report.engine.css.engine.value.FloatValue;
import org.eclipse.birt.report.engine.css.engine.value.css.CSSConstants;
import org.eclipse.birt.report.engine.emitter.xls.layout.ExcelContext;

public class POIWriterXLS
{

    private OutputStream out;
    private Workbook workbook;
    private Sheet currentSheet;
    private int currentRowNo = -1;
    private Row currentRow;
    private Cell currentCell;
    private ExcelContext context;
    private Map< Integer, CellStyle > styleIdCellStyleMap = null;
    protected static Logger logger = Logger.getLogger( POIWriterXLS.class.getName() );

    public POIWriterXLS()
    {
        styleIdCellStyleMap = new HashMap< Integer, CellStyle >();
    }

    public OutputStream getOutputStream()
    {
        return out;
    }

    public void open( OutputStream out, String string, ExcelContext context )
    {
        this.out = out;
        this.context = context;

        if ( context.getOfficeVersion().equals( "office2003" ) )
        {
            workbook = new HSSFWorkbook();
        }
        else
        {
            workbook = new XSSFWorkbook();
        }
    }

    public void open( OutputStream out, String string )
    {
        this.out = out;
        workbook = new HSSFWorkbook();

    }

    public void startSheet( String sheetname, double[] coordinates, boolean isRightToLeft )
    {
        currentRowNo = -1;

        currentSheet = workbook.createSheet( sheetname );
        for ( int col = 0; col < coordinates.length; col++ )
        {
            //Round the input values to 750th
            int real = (int)( coordinates[ col ] / 750 );
            int reminder = (int)( coordinates[ col ] % 750 );
            if ( ( reminder - 375 ) >= 0 )
            {
                real++;
            }
            int finalCoordinates = ( real * 750 );
            currentSheet.setColumnWidth( col, (int)Math.round( finalCoordinates * 0.048333 ) );
            // System.out.println( "Col : " + col + " -- Width : " + (int)Math.round( finalCoordinates * 0.048333 ) );
        }

        //currentSheet.setRightToLeft( isRightToLeft );

    }

    public void createRow( float rowHeight )
    {
        currentRowNo++;
        currentRow = currentSheet.createRow( currentRowNo );
        if ( rowHeight > 0 )
        {
            currentRow.setHeightInPoints( rowHeight );
        }
        else
        {
            //TODO: Need to set AutoFit Height
            logger.info( "Sheet :" + currentSheet.getSheetName() + " -- Unable to set the Autofit Row Height for Row: " + currentRowNo );
            //currentRow.setHeight( (short)-1 );
        }

    }

    public void createCell( int column, int colSpan, int rowSpan, int styleId, StyleEntry style, HyperlinkDef hyperLink, String urlAddress )
    {
        column = column - 1;
        //System.out.println( "Row : " + currentRowNo + " -- Column : " + column + " : colSpan " + colSpan + " : rowSpan " + rowSpan );
        currentCell = currentRow.createCell( column );
        CellStyle currentCellStyle = declareStyle( currentCell, style, styleId );
        if ( currentCellStyle != null )
        {
            currentCell.setCellStyle( currentCellStyle );
        }

        if ( hyperLink != null )
        {
            Hyperlink link = workbook.getCreationHelper().createHyperlink( Hyperlink.LINK_URL );
            link.setAddress( urlAddress );
            currentCell.setHyperlink( link );
            if ( hyperLink.getToolTip() != null )
            {
                //TODO: Need to set the tool tip info
            }
        }

        if ( colSpan != 0 || rowSpan != 0 )
        {
            //System.out.println( "Merging =  From Row " + currentRowNo + " : Last Row " + currentRowNo + rowSpan +" : Start Col " + column +" : End Col " +  column + colSpan  );
            CellRangeAddress mergedRegion = new CellRangeAddress( currentRowNo, currentRowNo + rowSpan, column, column + colSpan );
            currentSheet.addMergedRegion( mergedRegion );
            setStyleMergedRegion( mergedRegion, currentCellStyle );
        }

    }

    private void setStyleMergedRegion( CellRangeAddress mergedRegion, CellStyle currentCellStyle )
    {
        RegionUtil.setBorderBottom( currentCellStyle.getBorderBottom(), mergedRegion, currentSheet, workbook );
        RegionUtil.setBorderLeft( currentCellStyle.getBorderLeft(), mergedRegion, currentSheet, workbook );
        RegionUtil.setBorderRight( currentCellStyle.getBorderRight(), mergedRegion, currentSheet, workbook );
        RegionUtil.setBorderTop( currentCellStyle.getBorderTop(), mergedRegion, currentSheet, workbook );
        RegionUtil.setBottomBorderColor( currentCellStyle.getBottomBorderColor(), mergedRegion, currentSheet, workbook );
        RegionUtil.setLeftBorderColor( currentCellStyle.getLeftBorderColor(), mergedRegion, currentSheet, workbook );
        RegionUtil.setRightBorderColor( currentCellStyle.getRightBorderColor(), mergedRegion, currentSheet, workbook );
        RegionUtil.setTopBorderColor( currentCellStyle.getTopBorderColor(), mergedRegion, currentSheet, workbook );

    }

    public void writeText( int type, Object value, StyleEntry style )
    {

        String txt = ExcelUtil.format( value, type );
        if ( type == SheetData.NUMBER )
        {
            if ( ExcelUtil.isNaN( value ) || ExcelUtil.isBigNumber( value ) || ExcelUtil.isInfinity( value ) )
            {
                currentCell.setCellType( Cell.CELL_TYPE_STRING );
            }
            else
            {

                currentCell.setCellType( Cell.CELL_TYPE_NUMERIC );
                if ( value instanceof Integer )
                {
                    currentCell.setCellValue( (Integer)value );
                }
                else if ( value instanceof Long )
                {
                    currentCell.setCellValue( (Long)value );
                }
                else
                {
                    BigDecimal valueDecimal = new BigDecimal( value.toString() );
                    currentCell.setCellValue( valueDecimal.doubleValue() );
                }
                return;
            }
        }
        else if ( type == SheetData.DATE )
        {
            currentCell.setCellType( Cell.CELL_TYPE_NUMERIC );
            currentCell.setCellValue( (Date)value );
            return;
        }
        else
        {
            currentCell.setCellType( Cell.CELL_TYPE_STRING );
        }

        if ( style != null )
        {
            String textTransform = (String)style.getProperty( StyleConstant.TEXT_TRANSFORM );
            if ( CSSConstants.CSS_CAPITALIZE_VALUE.equalsIgnoreCase( textTransform ) )
            {
                txt = ExcelUtil.capitalize( txt );
            }
            else if ( CSSConstants.CSS_UPPERCASE_VALUE.equalsIgnoreCase( textTransform ) )
            {
                txt = txt.toUpperCase();
            }
            else if ( CSSConstants.CSS_LOWERCASE_VALUE.equalsIgnoreCase( textTransform ) )
            {
                txt = txt.toLowerCase();
            }
        }
        currentCell.setCellValue( ExcelUtil.truncateCellText( txt ) );

    }

    public void writeImage( int type, ImageData imageData, StyleEntry style, int column, int colSpan, int rowSpan )
    {
        column = column - 1;

        ClientAnchor anchor = null;
        int index = workbook.addPicture( imageData.getImageData(), Workbook.PICTURE_TYPE_JPEG );

        //TODO: Need to recheck the size of the image -- might require some change
        if ( workbook instanceof HSSFWorkbook )
        {
            anchor = new HSSFClientAnchor( 0, 0, 0, 0, (short)column, currentRowNo, (short)( column + colSpan + 1 ), currentRowNo + rowSpan );
        }
        else
        {
            anchor = new XSSFClientAnchor( 0, 0, 0, 0, column, currentRowNo, ( column + colSpan + 1 ), currentRowNo + rowSpan );
        }

        Drawing drawing = currentSheet.createDrawingPatriarch();
        drawing.createPicture( anchor, index );
    }

    private CellStyle declareStyle( Cell currentCell, StyleEntry style, int styleId )
    {

        CellStyle currentCellStyle = styleIdCellStyleMap.get( styleId );

        if ( currentCellStyle != null ) { return currentCellStyle; }

        if ( styleId >= StyleEngine.RESERVE_STYLE_ID )
        {

            currentCellStyle = currentCell.getSheet().getWorkbook().createCellStyle();

            //Set the Text Wrap Property
            boolean wrapText = context.getWrappingText();
            String whiteSpace = (String)style.getProperty( StyleConstant.WHITE_SPACE );
            if ( CSSConstants.CSS_NOWRAP_VALUE.equals( whiteSpace ) )
            {
                wrapText = false;
            }
            currentCellStyle.setWrapText( wrapText );

            // Horizontal Alignment
            setHorizontalAlignment( style, currentCellStyle );

            // Vertical Alignment
            setVerticalAlignment( style, currentCellStyle );

            // indent
            float indent = ExcelUtil.convertTextIndentToEM( (FloatValue)style.getProperty( StyleConstant.TEXT_INDENT ), (Float)style.getProperty( StyleConstant.FONT_SIZE_PROP ) );
            if ( indent != 0f )
            {
                // System.out.println("indent :"+indent);
                currentCellStyle.setIndention( (short)indent );
            }

            //TODO: Need to set the Direction
            /* String direction = (String) style.getProperty( StyleConstant.DIRECTION_PROP );
             if ( isValid( direction ) )
             {
                 if ( CSSConstants.CSS_RTL_VALUE.equals( direction ) )
                 {
                     
                 }
             }
             */

            //Bottom Border
            setBottomBorder( style, currentCellStyle );

            //Top Border
            setTopBorder( style, currentCellStyle );

            //Left Border
            setLeftBorder( style, currentCellStyle );

            //Right Border
            setRightBorder( style, currentCellStyle );

            //Set Font
            setFont( style, currentCellStyle );

            Color bgColor = (Color)style.getProperty( StyleConstant.BACKGROUND_COLOR_PROP );

            if ( bgColor != null )
            {
                if ( workbook instanceof HSSFWorkbook )
                {
                    currentCellStyle.setFillForegroundColor( getHSSFColorIndex( bgColor ) );
                }
                else
                {
                    ( (XSSFCellStyle)currentCellStyle ).setFillForegroundColor( getColor( bgColor ) );
                }
                currentCellStyle.setFillPattern( CellStyle.SOLID_FOREGROUND );
            }

        }

        String dataFormat = getDataFormat( style );

        if ( dataFormat != null )
        {

            if ( currentCellStyle == null )
            {
                currentCellStyle = currentCell.getSheet().getWorkbook().createCellStyle();
            }
            currentCellStyle.setDataFormat( workbook.createDataFormat().getFormat( dataFormat ) );
        }
        styleIdCellStyleMap.put( styleId, currentCellStyle );
        return currentCellStyle;

    }

    protected void setFont( StyleEntry style, CellStyle currentCellStyle )
    {
        String fontName = (String)style.getProperty( StyleConstant.FONT_FAMILY_PROP );
        Float size = (Float)style.getProperty( StyleConstant.FONT_SIZE_PROP );
        Boolean isItalic = (Boolean)style.getProperty( StyleConstant.FONT_STYLE_PROP );
        Boolean isBold = (Boolean)style.getProperty( StyleConstant.FONT_WEIGHT_PROP );
        Boolean strikeThrough = (Boolean)style.getProperty( StyleConstant.TEXT_LINE_THROUGH_PROP );
        Boolean underline = (Boolean)style.getProperty( StyleConstant.TEXT_UNDERLINE_PROP );
        Color color = (Color)style.getProperty( StyleConstant.COLOR_PROP );

        Font font = workbook.createFont();
        if ( isValid( fontName ) )
        {
            fontName = getFirstFont( fontName );
            font.setFontName( fontName );
        }
        //  System.out.println( "size :" + size );

        if ( size != null )
        {

            font.setFontHeightInPoints( size.shortValue() );
        }

        if ( strikeThrough != null && strikeThrough )
        {
            font.setStrikeout( strikeThrough );
        }

        if ( underline != null && underline )
        {
            font.setUnderline( (byte)1 );
        }

        if ( isBold )
        {
            
            font.setBoldweight( Font.BOLDWEIGHT_BOLD );
        }

        font.setItalic( isItalic );

        if ( color != null )
        {
            if ( workbook instanceof HSSFWorkbook )
            {
                font.setColor( getHSSFColorIndex( color ) );
            }
            else
            {
                ( (XSSFFont)font ).setColor( getColor( color ) );
            }

        }
        currentCellStyle.setFont( font );
    }

    protected void setRightBorder( StyleEntry style, CellStyle currentCellStyle )
    {
        //TODO: Need to set the line style
        Integer weight = (Integer)style.getProperty( StyleConstant.BORDER_RIGHT_WIDTH_PROP );

        if ( weight != null && weight > 0 )
        {
            switch (weight)
            {
            case 1:
                currentCellStyle.setBorderRight( CellStyle.BORDER_THIN );
                break;
            case 2:
                currentCellStyle.setBorderRight( CellStyle.BORDER_MEDIUM );
                break;
            case 3:
                currentCellStyle.setBorderRight( CellStyle.BORDER_THICK );
                break;
            default:
                currentCellStyle.setBorderRight( CellStyle.BORDER_THIN );

            }
        }
        Color color = (Color)style.getProperty( StyleConstant.BORDER_RIGHT_COLOR_PROP );
        if ( color != null )
        {
            if ( workbook instanceof HSSFWorkbook )
            {
                currentCellStyle.setRightBorderColor( getHSSFColorIndex( color ) );
            }
            else
            {
                ( (XSSFCellStyle)currentCellStyle ).setRightBorderColor( getColor( color ) );
            }

        }
    }

    protected void setLeftBorder( StyleEntry style, CellStyle currentCellStyle )
    {
        //TODO: Need to set the line style
        Integer weight = (Integer)style.getProperty( StyleConstant.BORDER_LEFT_WIDTH_PROP );

        if ( weight != null && weight > 0 )
        {
            switch (weight)
            {
            case 1:
                currentCellStyle.setBorderLeft( CellStyle.BORDER_THIN );
                break;
            case 2:
                currentCellStyle.setBorderLeft( CellStyle.BORDER_MEDIUM );
                break;
            case 3:
                currentCellStyle.setBorderLeft( CellStyle.BORDER_THICK );
                break;
            default:
                currentCellStyle.setBorderLeft( CellStyle.BORDER_THIN );

            }
        }

        Color color = (Color)style.getProperty( StyleConstant.BORDER_LEFT_COLOR_PROP );
        if ( color != null )
        {
            if ( workbook instanceof HSSFWorkbook )
            {
                currentCellStyle.setLeftBorderColor( getHSSFColorIndex( color ) );
            }
            else
            {
                ( (XSSFCellStyle)currentCellStyle ).setLeftBorderColor( getColor( color ) );
            }

        }
    }

    protected void setTopBorder( StyleEntry style, CellStyle currentCellStyle )
    {
        //TODO: Need to set the line style

        Integer weight = (Integer)style.getProperty( StyleConstant.BORDER_TOP_WIDTH_PROP );

        if ( weight != null && weight > 0 )
        {
            switch (weight)
            {
            case 1:
                currentCellStyle.setBorderTop( CellStyle.BORDER_THIN );
                break;
            case 2:
                currentCellStyle.setBorderTop( CellStyle.BORDER_MEDIUM );
                break;
            case 3:
                currentCellStyle.setBorderTop( CellStyle.BORDER_THICK );
                break;
            default:
                currentCellStyle.setBorderTop( CellStyle.BORDER_THIN );

            }
        }
        Color color = (Color)style.getProperty( StyleConstant.BORDER_TOP_COLOR_PROP );
        if ( color != null )
        {
            if ( workbook instanceof HSSFWorkbook )
            {
                currentCellStyle.setTopBorderColor( getHSSFColorIndex( color ) );
            }
            else
            {
                ( (XSSFCellStyle)currentCellStyle ).setTopBorderColor( getColor( color ) );
            }
        }
    }

    protected void setBottomBorder( StyleEntry style, CellStyle currentCellStyle )
    {
        //TODO: Need to set the line style
        //String bottomLineStyle = (String) style.getProperty( StyleConstant.BORDER_BOTTOM_STYLE_PROP );
        Integer weight = (Integer)style.getProperty( StyleConstant.BORDER_BOTTOM_WIDTH_PROP );

        if ( weight != null && weight > 0 )
        {
            switch (weight)
            {
            case 1:
                currentCellStyle.setBorderBottom( CellStyle.BORDER_THIN );
                break;
            case 2:
                currentCellStyle.setBorderBottom( CellStyle.BORDER_MEDIUM );
                break;
            case 3:
                currentCellStyle.setBorderBottom( CellStyle.BORDER_THICK );
                break;
            default:
                currentCellStyle.setBorderBottom( CellStyle.BORDER_THIN );
            }
        }

        Color color = (Color)style.getProperty( StyleConstant.BORDER_BOTTOM_COLOR_PROP );

        if ( color != null )
        {
            if ( workbook instanceof HSSFWorkbook )
            {
                currentCellStyle.setBottomBorderColor( getHSSFColorIndex( color ) );
            }
            else
            {
                ( (XSSFCellStyle)currentCellStyle ).setBottomBorderColor( getColor( color ) );
            }

        }
    }

    protected void setVerticalAlignment( StyleEntry style, CellStyle currentCellStyle )
    {
        String verticalAlign = (String)style.getProperty( StyleConstant.V_ALIGN_PROP );
        if ( isValid( verticalAlign ) )
        {
            // System.out.println("verticalAlign :"+verticalAlign);
            if ( verticalAlign.equals( "Bottom" ) )
            {
                currentCellStyle.setVerticalAlignment( CellStyle.VERTICAL_BOTTOM );
            }
            else if ( verticalAlign.equals( "Justify" ) )
            {
                currentCellStyle.setVerticalAlignment( CellStyle.VERTICAL_JUSTIFY );
            }
            else if ( verticalAlign.equals( "Center" ) )
            {
                currentCellStyle.setVerticalAlignment( CellStyle.VERTICAL_CENTER );
            }
            else if ( verticalAlign.equals( "Top" ) )
            {
                currentCellStyle.setVerticalAlignment( CellStyle.VERTICAL_TOP );
            }
        }
    }

    protected void setHorizontalAlignment( StyleEntry style, CellStyle currentCellStyle )
    {
        String horizontalAlign = (String)style.getProperty( StyleConstant.H_ALIGN_PROP );
        if ( isValid( horizontalAlign ) )
        {
            if ( horizontalAlign.equals( "Right" ) )
            {
                currentCellStyle.setAlignment( CellStyle.ALIGN_RIGHT );
            }
            else if ( horizontalAlign.equals( "Left" ) )
            {
                currentCellStyle.setAlignment( CellStyle.ALIGN_LEFT );
            }
            else if ( horizontalAlign.equals( "Center" ) )
            {
                currentCellStyle.setAlignment( CellStyle.ALIGN_CENTER );
            }
            else if ( horizontalAlign.equals( "Justify" ) )
            {
                currentCellStyle.setAlignment( CellStyle.ALIGN_JUSTIFY );
            }
        }
    }

    private short getHSSFColorIndex( Color color )
    {

        HSSFColor hssfColor = ( (HSSFWorkbook)workbook ).getCustomPalette().findSimilarColor( (byte)color.getRed(), (byte)color.getGreen(), (byte)color.getBlue() );
        if ( hssfColor != null )
        {
            return hssfColor.getIndex();
        }
        else
        {
            System.out.println( "Unable to find the color RBG: " + (byte)color.getRed() + "-" + (byte)color.getGreen() + "-" + (byte)color.getBlue() );
            return HSSFColor.BLACK.index;
        }

    }

    private XSSFColor getColor( Color color )
    {
        XSSFColor xssfColor = new XSSFColor( color );
        if ( xssfColor != null )
        {
            return xssfColor;
        }
        else
        {
            System.out.println( "Unable to find the color RBG: " + (byte)color.getRed() + "-" + (byte)color.getGreen() + "-" + (byte)color.getBlue() );
            return new XSSFColor();
        }
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

    private boolean isValid( String value )
    {
        return !StyleEntry.isNull( value );
    }

    public void close()
    {
        try
        {
            workbook.write( out );
            out.close();
        }
        catch ( IOException e )
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

    }

    private String getDataFormat( StyleEntry style )
    {
        Integer type = (Integer)style.getProperty( StyleConstant.DATA_TYPE_PROP );
        String dataFormat = null;
        if ( type == null )
            return dataFormat;
        if ( type == SheetData.DATE && style.getProperty( StyleConstant.DATE_FORMAT_PROP ) != null )
        {
            dataFormat = (String)style.getProperty( StyleConstant.DATE_FORMAT_PROP );
        }
        else if ( type == Data.NUMBER && style.getProperty( StyleConstant.NUMBER_FORMAT_PROP ) != null )
        {
            NumberFormatValue numberFormat = (NumberFormatValue)style.getProperty( StyleConstant.NUMBER_FORMAT_PROP );
            String format = numberFormat.getFormat();
            if ( format != null )
            {
                dataFormat = format;
            }
        }
        return dataFormat;
    }

    public void declareWorkSheetOptions( String orientation, int pageWidth, int pageHeight, float leftMargin, float rightMargin, float topMargin, float bottomMargin, String pageHeader, String pageFooter )
    {
        currentSheet.setDisplayGridlines( !context.getHideGridlines() );
        currentSheet.setMargin( Sheet.TopMargin, topMargin / ExcelUtil.INCH_PT );
        currentSheet.setMargin( Sheet.BottomMargin, bottomMargin / ExcelUtil.INCH_PT );
        currentSheet.setMargin( Sheet.LeftMargin, leftMargin / ExcelUtil.INCH_PT );
        currentSheet.setMargin( Sheet.RightMargin, rightMargin / ExcelUtil.INCH_PT );
        if ( pageHeader != null )
        {
            currentSheet.getHeader().setCenter( pageHeader );
        }
        if ( pageFooter != null )
        {
            currentSheet.getFooter().setCenter( pageFooter );
        }
        //TODO: we need to set the PaperSize information
    }

}
