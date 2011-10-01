package org.eclipse.birt.report.engine.emitter.xls;

import org.eclipse.birt.report.engine.emitter.xls.layout.Page;


public class RowData
{

	private SheetData[] rowdata;
	private double height;

	public RowData( Page page, SheetData[] rowdata, double height )
	{
		this.rowdata = rowdata;
		this.height = height;
	}

	public SheetData[] getRowdata( )
	{
		return rowdata;
	}

	public void setRowdata( SheetData[] rowdata )
	{
		this.rowdata = rowdata;
	}

	public double getHeight( )
	{
		return height;
	}

	public void setHeight( double height )
	{
		this.height = height;
	}
}
