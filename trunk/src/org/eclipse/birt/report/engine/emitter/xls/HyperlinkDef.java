package org.eclipse.birt.report.engine.emitter.xls;

import java.awt.Color;
import java.io.Serializable;

public class HyperlinkDef implements Serializable
{
	private static final long serialVersionUID = 5933271313761755249L;
	private String url;
	private int type;
    private String toolTip;
	private Color color;
    
	public HyperlinkDef( String url, int type, String toolTip )
	{
		this.url = url;
		this.type = type;
		this.toolTip = toolTip;
	}

	public String getUrl( )
	{
		return url;
	}
    
	public int getType( )
	{
		return type;
	}
	
	public void setUrl( String url) 
	{
	   	this.url = url;
	}
	
	public String getToolTip()
	{
		return toolTip;
	}
	
	public void setToolTip(String toolTip)
	{
		this.toolTip = toolTip;
	}

	public void setColor( Color color )
	{
		this.color = color;
	}

	public Color getColor( )
	{
		return this.color;
	}
}
