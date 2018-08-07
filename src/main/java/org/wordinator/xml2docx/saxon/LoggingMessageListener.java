package org.wordinator.xml2docx.saxon;

import javax.xml.transform.SourceLocator;

import org.apache.logging.log4j.Logger;

import net.sf.saxon.s9api.MessageListener;
import net.sf.saxon.s9api.XdmNode;

/**
 * Put Saxon xsl:message output to a log4j log
 *
 */
public class LoggingMessageListener implements MessageListener {

	private Logger log;

	public LoggingMessageListener(Logger log) {
		this.log = log;
	}

	@Override
	public void message(XdmNode content, boolean terminate, SourceLocator locator) {		
		String message = content.toString() + " (" + locator.getSystemId() + ") [" + locator.getLineNumber() + ":" + locator.getColumnNumber() + "]";
		if (message.contains("[ERR")) {
			log.error(message);
		} else if (message.contains("[WARN")) {
			log.warn(message);
		} else {
			log.info(message);
		}

	}

}
