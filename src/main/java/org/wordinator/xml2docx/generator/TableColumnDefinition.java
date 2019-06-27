package org.wordinator.xml2docx.generator;

import org.apache.poi.xwpf.usermodel.XWPFTable;

/**
 * Holds the properties for a single table column
 *
 */
public class TableColumnDefinition {


	private final TableColumnDefinitions parent;
	private int colIndex = -1;
	private String width = "auto";
	private String specifiedWidth = width;

	public TableColumnDefinition(TableColumnDefinitions tableColumnDefinitions) {
		this.parent = tableColumnDefinitions;
		this.colIndex = tableColumnDefinitions.getColumnDefinitions().size();
	}

	/**
	 * Set the width of the column
	 * @param width Width value, one of "auto", a measurement, or a percentage.
	 * @throws MeasurementException 
	 */
	public void setWidth(String width, int dotsPerInch) throws MeasurementException {
		this.specifiedWidth = width;
		this.width = interpretWidthSpecification(width, dotsPerInch);
	}

	/**
	 * Get the width value as appropriate for setting on XWPF widths
	 * @return The width value converted to "auto", twips, or a percentage.
	 */
	public String getWidth() {
		return width;
	}

	/**
	 * Get the width value as specified to setWidth()
	 * @return The specified widthvalue
	 */
	public String getSpecifiedWidth() {
		return specifiedWidth;
	}
	
	/**
	 * Get the column index for the column
	 * @return The zero-index position of the column definition in the list of column definitions
	 */
	public int getColumnIndex() {
		return this.colIndex;
	}
	
	/**
	 * Get the containing TableColumnDefinitions object
	 * @return Parent object
	 */
	public TableColumnDefinitions getParent() {
		return this.parent;
	}

	/**
	 * Set the width to "auto"
	 */
	public void setWidthAuto() {
		this.width = "auto";
		this.specifiedWidth = "auto";
		
	}

	/**
	 * Interpret a SWPX width specifier into an XWPF table width value.
	 * @param widthSpec Width specification ("auto", percentage, or measurement)
	 * @param dotsPerInch Dots per inch to use in converting measurements to twips
	 * @return Measurement value appropriate for use on XWPFTable.setWidth()
	 * @throws MeasurementException
	 */
	public static String interpretWidthSpecification(
								String widthSpec, 
								int dotsPerInch) 
			throws MeasurementException {
		if (widthSpec.matches(XWPFTable.REGEX_WIDTH_VALUE)) {
			return widthSpec;			
		} else {
			// Must be a measurement
		  long twips = Measurement.toTwips(widthSpec, dotsPerInch);
			return Long.toString(twips);
		}

	}

}
