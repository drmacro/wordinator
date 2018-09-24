package org.wordinator.xml2docx.generator;

/**
 * Utilities for working with measurements.
 *
 */
public class Measurement {
	
	public final static int POINTS_PER_INCH = 72;
	private static final int POINTS_PER_PICA = 12;

	/**
	 * Calculate the number of pixels represented by the specified value.
	 * @param measurementValue The measurement value with unit (e.g., "12pt", "1.3in")
	 * @param dotsPerInch The number of pixels (dots) per inch.s
	 * @return Number of pixels
	 * @throws MeasurementException Bad measurement value
	 */
	public static double toPixels(String measurementValue, int dotsPerInch) throws MeasurementException {
		String value = measurementValue.toLowerCase();
		Double result = 0.0;
		try {
			if (value.endsWith("px") || value.matches("\\-?[0-9\\.]+")) {
				String numberStr = value.endsWith("px") ? value.substring(0, value.length() - 2) : value;
				result = Double.parseDouble(numberStr);
			} else {
				double inches = toInches(measurementValue, dotsPerInch);
				result = inches * dotsPerInch;				
			}
		} catch (NumberFormatException e) {
			throw new MeasurementException("Measurement '" + value + "' is not numeric.");
		}
		
		return result;
						
	}

	/**
	 * Calculate the number of points represented by the specified measurement value.
	 * @param measurementValue The measurement value with unit (e.g., "12pt", "1.3in")
	 * @param dotsPerInch The number of pixels (dots) per inch.
	 * @return Number of points
	 * @throws MeasurementException Bad measurement value 
	 */
	public static double toPoints(String measurementValue, int dotsPerInch) throws MeasurementException {
		Double result = 0.0;
		double inches = toInches(measurementValue, dotsPerInch);
		result = inches * POINTS_PER_INCH;				
		
		return result;
	}

	/**
	 * Calculate the number of inches represented by the specified measurement value.
	 * @param measurementValue The measurement value with unit (e.g., "12pt", "1.3in")
	 * @param dotsPerInch The number of pixels (dots) per inch.
	 * @return Number of inchines
	 * @throws MeasurementException Bad measurement value
	 */
	public static double toInches(String measurementValue, int dotsPerInch)
			throws MeasurementException {
		String value = measurementValue.toLowerCase();
		String numberStr = value.substring(0, value.length() - 2);
		double inches;
		if (value.endsWith("pt")) {
			double points = Double.parseDouble(numberStr);
			inches = points / POINTS_PER_INCH;
		} else if (value.endsWith("pc")) {
			Double picas = Double.parseDouble(numberStr);
			inches = (picas * POINTS_PER_PICA) / POINTS_PER_INCH;
		} else if (value.endsWith("px")) {
			Double pixels = Double.parseDouble(numberStr);
			inches = pixels / dotsPerInch;
		} else if ( value.matches("\\-?[0-9\\.]+")) {
			Double pixels = Double.parseDouble(value);
			inches = pixels / dotsPerInch;
		} else if (value.endsWith("in")) {
			inches = Double.parseDouble(numberStr);
		} else if (value.endsWith("mm")) {
			double mms = Double.parseDouble(numberStr);
			inches = mms / 25.4;
		} else if (value.endsWith("cm")) {
			double cms = Double.parseDouble(numberStr);
			inches = cms / 2.54;
		} else {
			throw new UnrecognizedUnitException("Unrecognized unit for measurement '" + measurementValue + "'");
		}
		return inches;
	}

	/**
	 * Calculate the number twips (1/20th of a point) represented by the measurement value.
	 * @param measurementValue The measurement value with unit (e.g., "12pt", "1.3in") 
	 * @param dotsPerInch The number of pixels (dots) per inch.
	 * @return Number of twips
	 * @throws MeasurementException Bad measurement value 
	 */
	public static int toTwips(String measurementValue, int dotsPerInch) throws MeasurementException {
		double points = toPoints(measurementValue, dotsPerInch);
		return (int) points * 20;
	}

}
