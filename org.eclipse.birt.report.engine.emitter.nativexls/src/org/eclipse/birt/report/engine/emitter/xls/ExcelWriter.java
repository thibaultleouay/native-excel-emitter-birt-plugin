package org.eclipse.birt.report.engine.emitter.xls;

import java.io.*;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

import org.eclipse.birt.report.engine.content.IReportContent;
import org.eclipse.birt.report.engine.emitter.xls.layout.ExcelContext;

public class ExcelWriter implements IExcelWriter
{

    private NativeExcelWriter writer, tempWriter;
    private final ExcelContext context;
    private String tempFilePath;
    private int sheetIndex = 1;

    /**
     * @param out
     * @param context
     * @param isRtlSheet
     * @param pageHeader
     * @param pageFooter
     * @param orientation
     */
    public ExcelWriter( ExcelContext context )
    {
        this.context = context;
    }


    public void end( ) throws IOException
    {
        writer.end( );
		if ( tempFilePath != null )
		{
			File file = new File( tempFilePath );
			if ( file.exists( ) && file.isFile( ) )
			{
				file.delete( );
			}
		}
    }

    public void endRow( )
    {
        writer.endRow( );
    }

    public void endSheet( double[] coordinates, String oritentation,
            int pageWidth, int pageHeight, float leftMargin, float rightMargin,
            float topMargin, float bottomMargin )
    {
        writer.endSheet( coordinates, oritentation, pageWidth, pageHeight,
                leftMargin, rightMargin, topMargin, bottomMargin );
    }

    public void outputData( String sheet, SheetData data, StyleEntry style, int column,
            int colSpan ) throws IOException
    {
        //TODO: ignored sheet temporarily
        outputData( data, style, column, colSpan );
    }

    public void outputData( SheetData data, StyleEntry style, int column,
            int colSpan ) throws IOException
    {
        writer.outputData( data, style, column, colSpan );
    }

    public void start( IReportContent report, Map<StyleEntry, Integer> styles,
    // TODO: style ranges.
            // List<ExcelRange> styleRanges,
            HashMap<String, BookmarkDef> bookmarkList ) throws IOException
    {
        writer = new NativeExcelWriter( context );
        writer.setSheetIndex( sheetIndex );
        // TODO: style ranges.
        // writer.start( report, styles, styleRanges, bookmarkList );
        writer.start( report, styles, bookmarkList );
        copyOutputData( );
    }

    private void copyOutputData( ) throws IOException
    {
        if ( tempWriter != null )
        {
            tempWriter.close( );
            FileInputStream fis = null;
            OutputStream outputStream = writer.getWriter().getOutputStream();
            try
            {
                fis = new FileInputStream( tempFilePath );
                
                byte[] buf = new byte[1024];
                int i = 0;
                while ( (i = fis.read( buf )) != -1 )
                {
                    outputStream.write( buf, 0, i );
                }
            }
            finally
            {
                fis.close();
                outputStream.close();
            }
        }
        
    }

    public void startRow( double rowHeight )
    {
        writer.startRow( rowHeight );
    }

    public void startSheet( String name ) throws IOException
    {
        if ( writer == null )
        {
            initializeWriterAsTempWriter( );
        }
        writer.startSheet( name );
        sheetIndex++;
    }

    public void startSheet( double[] coordinates, String pageHeader,
            String pageFooter, String name ) throws IOException
    {
        if ( writer == null )
        {
            initializeWriterAsTempWriter( );
        }
        writer.startSheet( coordinates, pageHeader, pageFooter, name );
        sheetIndex++;
    }

    /**
     * @throws FileNotFoundException
     * 
     */
    private void initializeWriterAsTempWriter( ) throws FileNotFoundException
    {
        String tempFolder = context.getTempFileDir( );
        if ( !( tempFolder.endsWith( "/" ) || tempFolder.endsWith( "\\" ) ) )
        {
            tempFolder = tempFolder.concat( "/" );
        }
        tempFilePath = tempFolder + "birt_xls_"
                + UUID.randomUUID( ).toString( );
        FileOutputStream out = new FileOutputStream( tempFilePath );
        tempWriter = new NativeExcelWriter( out, context );
        writer = tempWriter;
    }

    public void endSheet( )
    {
        writer.endSheet( );
    }

    public void startRow( )
    {
        writer.startRow( );
    }

    public void outputData( int col, int row, int type, Object value )
    {
        writer.outputData( col, row, type, value );
    }

    public String defineName( String cells )
    {
        return writer.defineName( cells );
    }
}
