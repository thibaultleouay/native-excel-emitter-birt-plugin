package org.eclipse.birt.report.engine.emitter.xls.layout;

import org.eclipse.birt.report.engine.emitter.xls.StyleEntry;


public class XlsCell extends XlsContainer
{
	private int rowSpan;
	
	public XlsCell( StyleEntry style, ContainerSizeInfo sizeInfo,
			XlsContainer parent, int rowSpan )
	{
		super( style, sizeInfo, parent );
		this.rowSpan = rowSpan;
	}

	public int getRowSpan( )
	{
		return rowSpan;
	}
}
