package org.wordinator.xml2docx.saxon;

import javax.xml.transform.stream.StreamResult;

import net.sf.saxon.lib.Logger;

/**
 * Log saxon messages to Log4J logger
 *
 */
public class Log4jSaxonLogger extends Logger {

	private org.apache.logging.log4j.Logger log;

	public Log4jSaxonLogger(org.apache.logging.log4j.Logger log) {
		this.log = log;
	}

	@Override
	public void println(String message, int severity) {
		switch (severity) {
		case WARNING:
			log.warn(message);
			break;
		case ERROR:
			log.error(message);
			break;
		case DISASTER:
			log.fatal(message);
			break;
		case INFO:
		default:
			log.info(message);		
		}
	}

	@Override
	public StreamResult asStreamResult() {
		// TODO Auto-generated method stub
		return null;
	}

}
