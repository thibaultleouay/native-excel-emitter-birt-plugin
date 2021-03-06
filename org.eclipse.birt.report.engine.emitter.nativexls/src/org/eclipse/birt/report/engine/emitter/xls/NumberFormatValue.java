package org.eclipse.birt.report.engine.emitter.xls;

import java.math.RoundingMode;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 
 */

public class NumberFormatValue
{

	private int fractionDigits;

	private String format;
	private RoundingMode roundingMode;
	private static Pattern pattern = Pattern.compile( "^(.*?)\\{RoundingMode=(.*?)\\}",
			Pattern.CASE_INSENSITIVE );

	private NumberFormatValue( )
	{
	}

	public static NumberFormatValue getInstance( String numberFormat )
	{
		if ( numberFormat != null )
		{
			NumberFormatValue value = new NumberFormatValue( );
			Matcher matcher = pattern.matcher( numberFormat );
			if ( matcher.matches( ) )
			{
				String f = matcher.group( 1 );
				if ( f != null && f.length( ) > 0 )
				{
					value.format = f;
					int index = f.lastIndexOf( '.' );
					if ( index > 0 )
					{
						int end = f.length( );
						for ( int i = index + 1; i < f.length( ); i++ )
						{
							if ( f.charAt( i ) != '0' )
							{
								end = i;
								break;
							}
						}
						value.fractionDigits = end - 1 - index;
					}
					char lastChar = f.charAt( f.length( ) - 1 );
					switch ( lastChar )
					{
						case 37 :
							value.fractionDigits += 2;
							break;
						case 8240: 
							value.fractionDigits += 3;
							break;
						case 8241: 
							value.fractionDigits += 4;
							break;
					}
				}
				String m = matcher.group( 2 );
				if ( m != null )
				{
					value.roundingMode = RoundingMode.valueOf( m );
				}
			}
			else
			{
				value.format = numberFormat;
			}
			return value;
		}
		return null;
	}

	public int getFractionDigits( )
	{
		return fractionDigits;
	}

	public void setFractionDigits( int fractionDigits )
	{
		this.fractionDigits = fractionDigits;
	}

	public String getFormat( )
	{
		return format;
	}

	public void setFormat( String format )
	{
		this.format = format;
	}

	public RoundingMode getRoundingMode( )
	{
		return roundingMode;
	}

	public void setRoundingMode( RoundingMode roundingMode )
	{
		this.roundingMode = roundingMode;
	}

}
