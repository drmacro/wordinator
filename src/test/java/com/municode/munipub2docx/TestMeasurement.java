package com.municode.munipub2docx;

import org.junit.Test;

import com.municode.munipub2docx.generator.Measurement;
import com.municode.munipub2docx.generator.UnrecognizedUnitException;

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
		String measureInches = "10in"; // Inches
		String measureMM = "100mm"; // Millimeters
		String measureCM = "10cm"; // Centimeters
		String measureBogus = "100bg"; // Not a real unit
		
		int dotsPerInch = 72;
		
		assertEquals("Bare pixel failed.", 100.0, Measurement.toPixels(measurePixelsBare, dotsPerInch ));
		assertEquals("PX unit failed.", 100.0, Measurement.toPixels(measurePixelsPX, dotsPerInch ));
		assertEquals("PT unit failed.", 100.0, Measurement.toPixels(measurePoints, dotsPerInch ));
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
		String measureInches = "10in"; // Inches
		String measureMM = "100mm"; // Millimeters
		String measureCM = "10cm"; // Centimeters
		String measureBogus = "100bg"; // Not a real unit
		
		int dotsPerInch = 72;
		
		assertEquals("Bare pixel failed.", 100.0, 
				Measurement.toPoints(measurePixelsBare, dotsPerInch ));
		assertEquals("PX unit failed.", 100.0, Measurement.toPoints(measurePixelsPX, dotsPerInch ));
		assertEquals("PT unit failed.", 100.0, Measurement.toPoints(measurePoints, dotsPerInch ));
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
	public void testToTwips() throws Throwable {
		String measurePixelsBare = "100"; // Pixels
		String measurePixelsPX = "100px"; // Pixels
		String measurePoints = "100pt"; // Points
		String measureInches = "10in"; // Inches
		String measureMM = "100mm"; // Millimeters
		String measureCM = "10cm"; // Centimeters
		String measureBogus = "100bg"; // Not a real unit
		
		int dotsPerInch = 72;
		
		assertEquals("Bare pixel failed.", 100 * 20, 
				Measurement.toTwips(measurePixelsBare, dotsPerInch ));
		assertEquals("PX unit failed.", 100 * 20, Measurement.toTwips(measurePixelsPX, dotsPerInch ));
		assertEquals("PT unit failed.", 100 * 20, Measurement.toTwips(measurePoints, dotsPerInch ));
		assertEquals("IN unit failed.", 10 * 20 * 72, Measurement.toTwips(measureInches, dotsPerInch ));
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
