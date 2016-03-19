package org.eclipse.birt.report.engine.emitter.xls;




public class Data extends SheetData
{

	public Data( )
	{
	}

	public Data( SheetData data )
	{
		this( data.getValue( ), data.getStyleId( ), data.getDataType( ) );
		this.rowIndex = data.getRowIndex( );
	}
	
	public Data( final Object value, final int datatype )
	{
		this( value, 0, datatype );
	}

	public Data( final Object value, final int styleId, final int datatype )
	{
		this( value, styleId, datatype, 0 );
	}

	public Data( final Object value, final int styleId, final int datatype,
			int rowSpanOfDesign )
	{
		this.value = value;
		this.styleId = styleId;
		this.dataType = datatype;
		this.rowSpanInDesign = rowSpanOfDesign;
	}

	public Object getValue( )
	{
		return value;
	}
}