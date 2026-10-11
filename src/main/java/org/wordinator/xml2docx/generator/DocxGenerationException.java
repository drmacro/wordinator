package org.wordinator.xml2docx.generator;

/*
 * Reports problems generating DOCX files.
 */
public class DocxGenerationException extends Exception {

  public DocxGenerationException(String string) {
    super(string);
  }

  public DocxGenerationException(Throwable cause) {
    super(cause);
  }

  public DocxGenerationException(String string, Throwable cause) {
    super(string, cause);
  }

  public DocxGenerationException(String message, Throwable cause,
                                 boolean enableSuppression,
                                 boolean writableStackTrace) {
    super(message, cause, enableSuppression, writableStackTrace);
  }

}
