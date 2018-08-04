package com.municode.munipub2docx.generator;

import java.io.File;
import java.net.URL;

import javax.xml.transform.Result;
import javax.xml.transform.TransformerException;
import javax.xml.transform.sax.SAXResult;

import org.apache.commons.io.FilenameUtils;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlSaxHandler;

import net.sf.saxon.lib.OutputURIResolver;

/**
 * Saxon S9 OutputURIResolver implementation that takes the result and generates
 * a DOCX file from it.
 *
 */
public class DocxGeneratingOutputUriResolver implements OutputURIResolver {

	private File outDir;
	private File templateFile;
	private XmlSaxHandler saxHandler;

	/**
	 * 
	 * @param outDir Directory to put new DOCX files into.
	 * @param templateFile The DOTX template to use in constructing new DOCX files.
	 */
	public DocxGeneratingOutputUriResolver(File outDir, File templateFile) {
		this.outDir = outDir;
		this.templateFile = templateFile;
	}

	public OutputURIResolver newInstance() {
		return new DocxGeneratingOutputUriResolver(outDir, templateFile);
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
			String filename = FilenameUtils.getBaseName(result.getSystemId()) + ".docx";
			File outFile = new File(outDir, filename);
			File inFile = new File(new URL(result.getSystemId()).toURI());
			System.out.println("+ [INFO] Generating DOCX file \"" + outFile.getAbsolutePath() + "\"");
			DocxGenerator generator = new DocxGenerator(inFile, outFile, templateFile);
			generator.generate(xml);
		} catch (Exception e) {
			throw new TransformerException(e);
		}

	}

}
