package com.municode.munipub2docx.generator;

/*
 * Reports problems generating DOCX files.
 */
public class DocxGenerationException extends Exception {

	private static final long serialVersionUID = 1L;

	public DocxGenerationException(String string) {
		super(string);
	}
	
	public DocxGenerationException(String string, Throwable cause) {
		super(string, cause);
	}

}
