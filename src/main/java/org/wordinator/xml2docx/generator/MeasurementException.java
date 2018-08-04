package org.wordinator.xml2docx.generator;

public class MeasurementException extends Exception {

	private static final long serialVersionUID = 1L;

	public MeasurementException() {
	}

	public MeasurementException(String message) {
		super(message);
	}

	public MeasurementException(Throwable cause) {
		super(cause);
	}

	public MeasurementException(String message, Throwable cause) {
		super(message, cause);
	}

	public MeasurementException(String message, Throwable cause, boolean enableSuppression,
			boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}

}
