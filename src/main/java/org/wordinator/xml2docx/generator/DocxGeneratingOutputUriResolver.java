package org.wordinator.xml2docx.generator;

import java.io.File;
import java.net.URL;
import java.net.URLDecoder;

import javax.xml.transform.Result;
import javax.xml.transform.TransformerException;
import javax.xml.transform.sax.SAXResult;

import org.apache.commons.io.FilenameUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlSaxHandler;

import net.sf.saxon.lib.OutputURIResolver;

/**
 * Saxon S9 OutputURIResolver implementation that takes the result and generates
 * a DOCX file from it.
 *
 */
public class DocxGeneratingOutputUriResolver implements OutputURIResolver {
	
	public static Logger log = LogManager.getLogger();

	private File outDir;
	private File templateFile;
	private XmlSaxHandler saxHandler;

	private int dotsPerInch = 96; // FIXME: Need to figure out a way to make this
	                              // configurable given that resolver is created using
								  // newInstance()

	/**
	 * 
	 * @param outDir Directory to put new DOCX files into.
	 * @param templateFile The DOTX template to use in constructing new DOCX files.
	 * @param log 
	 */
	public DocxGeneratingOutputUriResolver(File outDir, File templateFile, Logger log) {
		this.outDir = outDir;
		this.templateFile = templateFile;
		DocxGeneratingOutputUriResolver.log = log;
	}

	public OutputURIResolver newInstance() {
		return new DocxGeneratingOutputUriResolver(outDir, templateFile, log);
	}

	public Result resolve(String href, String base) throws TransformerException {
		saxHandler = XmlObject.Factory.newXmlSaxHandler();
	
		Result result = new SAXResult(saxHandler.getContentHandler());
		result.setSystemId(href);
		return result;
		
	}

	public void close(Result result) throws TransformerException {
		// Do the DOCX building
		
		try {
			XmlObject xml = saxHandler.getObject();
			String outFilepath = URLDecoder.decode(result.getSystemId(), "UTF-8");
			String filename = FilenameUtils.getBaseName(outFilepath) + ".docx";
			File outFile = new File(outDir, filename);
			File inFile = new File(new URL(result.getSystemId()).toURI());
			log.info("Generating DOCX file \"" + outFile.getAbsolutePath() + "\"");
			DocxGenerator generator = new DocxGenerator(inFile, outFile, templateFile);
			generator.setDotsPerInch(dotsPerInch);
			generator.generate(xml);
		} catch (Exception e) {
			throw new TransformerException(e);
		}

	}

	public int getDotsPerInch() {
		return dotsPerInch;
	}

}
