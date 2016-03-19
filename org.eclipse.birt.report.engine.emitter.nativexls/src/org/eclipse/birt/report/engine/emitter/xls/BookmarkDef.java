package org.eclipse.birt.report.engine.emitter.xls;

public class BookmarkDef
{

	private String name;
	private int columnNo;
	private int rowNo;
	private int sheetIndex;
	private String generatedName;

	public String getValidName( )
	{
		return generatedName == null ? name : generatedName;
	}

	public void setGeneratedName( String generatedName )
	{
		this.generatedName = generatedName;
	}

	public BookmarkDef( String name )
	{
		this.name = name;
	}

	public String getName( )
	{
		return name;
	}

	public void setName( String name )
	{
		this.name = name;
	}

	public int getColumnNo( )
	{
		return columnNo;
	}

	public void setColumnNo( int columnNo )
	{
		this.columnNo = columnNo;
	}

	public int getRowNo( )
	{
		return rowNo;
	}

	public void setRowNo( int rowNo )
	{
		this.rowNo = rowNo;
	}

	public int getSheetIndex( )
	{
		return this.sheetIndex;
	}

	public void setSheetIndex( int sheetIndex )
	{
		this.sheetIndex = sheetIndex;
	}
}
