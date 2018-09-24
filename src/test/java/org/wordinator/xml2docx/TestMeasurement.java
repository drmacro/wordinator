package org.wordinator.xml2docx;

import org.junit.Test;
import org.wordinator.xml2docx.generator.Measurement;
import org.wordinator.xml2docx.generator.UnrecognizedUnitException;

import junit.framework.TestCase;

/**
 * Tests for the Measurement utility class.
 *
 */
public class TestMeasurement extends TestCase {
	
	@Test
	public void testToPixels() throws Throwable {
		String measurePixelsBare = "100"; // Pixels
		String measurePixelsPX = "100px"; // Pixels
		String measurePoints = "100pt"; // Points
		String measurePicas = "10pc"; // Picas
		String measureInches = "10in"; // Inches
		String measureMM = "100mm"; // Millimeters
		String measureCM = "10cm"; // Centimeters
		String measureBogus = "100bg"; // Not a real unit
		
		int dotsPerInch = 72;
		
		assertEquals("Bare pixel failed.", 100.0, Measurement.toPixels(measurePixelsBare, dotsPerInch ));
		assertEquals("PX unit failed.", 100.0, Measurement.toPixels(measurePixelsPX, dotsPerInch ));
		assertEquals("PT unit failed.", 100.0, Measurement.toPixels(measurePoints, dotsPerInch ));
		assertEquals("PC unit failed.", 120.0, Measurement.toPixels(measurePicas, dotsPerInch ));
		assertEquals("IN unit failed.", 10.0 * dotsPerInch, Measurement.toPixels(measureInches, dotsPerInch ));
		assertEquals("MM unit failed.", "283.465", 
					 String.format("%.3f", Measurement.toPixels(measureMM, dotsPerInch ))
				    );
		assertEquals("CM unit failed.", "283.465", 
				     String.format("%.3f", Measurement.toPixels(measureCM, dotsPerInch ))
				    );
		try {
			Measurement.toPixels(measureBogus, dotsPerInch);
			fail("Bogus measurement " + measureBogus + " did not cause exception.");
		} catch (UnrecognizedUnitException e) {
			// Expected exception.
		}
	}
	
	@Test
	public void testToPoints() throws Throwable {
		String measurePixelsBare = "100"; // Pixels
		String measurePixelsPX = "100px"; // Pixels
		String measurePoints = "100pt"; // Points
		String measurePicas = "10pc"; // Picas
		String measureInches = "10in"; // Inches
		String measureMM = "100mm"; // Millimeters
		String measureCM = "10cm"; // Centimeters
		String measureBogus = "100bg"; // Not a real unit
		
		int dotsPerInch = 72;
		
		assertEquals("Bare pixel failed.", 100.0, 
				Measurement.toPoints(measurePixelsBare, dotsPerInch ));
		assertEquals("PX unit failed.", 100.0, Measurement.toPoints(measurePixelsPX, dotsPerInch ));
		assertEquals("PT unit failed.", 100.0, Measurement.toPoints(measurePoints, dotsPerInch ));
		assertEquals("PC unit failed.", 120.0, Measurement.toPoints(measurePicas, dotsPerInch ), 0.01);
		assertEquals("IN unit failed.", 10.0 * 72, Measurement.toPoints(measureInches, dotsPerInch ));
		assertEquals("MM unit failed.", "283.465", 
					 String.format("%.3f", Measurement.toPoints(measureMM, dotsPerInch ))
				    );
		assertEquals("CM unit failed.", "283.465", 
				     String.format("%.3f", Measurement.toPoints(measureCM, dotsPerInch ))
				    );
		try {
			Measurement.toPoints(measureBogus, dotsPerInch);
			fail("Bogus measurement " + measureBogus + " did not cause exception.");
		} catch (UnrecognizedUnitException e) {
			// Expected exception.
		}
	}
	
	@Test
	public void testToInches() throws Throwable {
		String measurePixelsBare = "100"; // Pixels
		String measurePixelsPX = "100px"; // Pixels
		String measurePoints = "100pt"; // Points
		String measurePicas = "10pc"; // Picas
		String measureInches = "10in"; // Inches
		String measureMM = "100mm"; // Millimeters
		String measureCM = "10cm"; // Centimeters
		String measureBogus = "100bg"; // Not a real unit
		
		int dotsPerInch = 72;
		
		assertEquals("Bare pixel failed.", 1.38, 
				Measurement.toInches(measurePixelsBare, dotsPerInch ), 0.1);
		assertEquals("PX unit failed.", 1.38, Measurement.toInches(measurePixelsPX, dotsPerInch ), 0.1);
		assertEquals("PT unit failed.", 1.38, Measurement.toInches(measurePoints, dotsPerInch ), 0.1);
		assertEquals("PC unit failed.", 10.0 / 6, Measurement.toInches(measurePicas, dotsPerInch ), 0.1);
		assertEquals("IN unit failed.", 10.0, Measurement.toInches(measureInches, dotsPerInch ), 0.1);
		assertEquals("MM unit failed.", 3.94, Measurement.toInches(measureMM, dotsPerInch ), 0.1);
		assertEquals("CM unit failed.", 3.94, Measurement.toInches(measureCM, dotsPerInch ), 0.1);
		try {
			Measurement.toInches(measureBogus, dotsPerInch);
			fail("Bogus measurement " + measureBogus + " did not cause exception.");
		} catch (UnrecognizedUnitException e) {
			// Expected exception.
		}
	}
	
	@Test
	public void testToTwips() throws Throwable {
		String measurePixelsBare = "100"; // Pixels
		String measurePixelsPX = "100px"; // Pixels
		String measurePoints = "100pt"; // Points
		String measurePicas = "10pc"; // Picas
		String measureInches = "10in"; // Inches
		String measureMM = "100mm"; // Millimeters
		String measureCM = "10cm"; // Centimeters
		String measureBogus = "100bg"; // Not a real unit
		
		int dotsPerInch = 72;
		
		assertEquals("Bare pixel failed.", 100 * 20, 
				Measurement.toTwips(measurePixelsBare, dotsPerInch ));
		assertEquals("PX unit failed.", 100 * 20, Measurement.toTwips(measurePixelsPX, dotsPerInch ));
		assertEquals("PT unit failed.", 100 * 20, Measurement.toTwips(measurePoints, dotsPerInch ));
		assertEquals("PC unit failed.", 10 * 12 * 20, Measurement.toTwips(measurePicas, dotsPerInch ), 0.1);
		assertEquals("IN unit failed.", 10 * 72 * 20, Measurement.toTwips(measureInches, dotsPerInch ));
		assertEquals("MM unit failed.", 5660, Measurement.toTwips(measureMM, dotsPerInch));
		assertEquals("CM unit failed.", 5660, Measurement.toTwips(measureCM, dotsPerInch));
		try {
			Measurement.toTwips(measureBogus, dotsPerInch);
			fail("Bogus measurement " + measureBogus + " did not cause exception.");
		} catch (UnrecognizedUnitException e) {
			// Expected exception.
		}
	}
	
}
