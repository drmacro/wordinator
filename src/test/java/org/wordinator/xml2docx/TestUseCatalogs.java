package org.wordinator.xml2docx;

import java.io.File;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.Options;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.junit.Test;

import junit.framework.TestCase;

public class TestUseCatalogs extends TestCase {

  private static final String DOTX_TEMPLATE_PATH = "docx/Test_Template.dotx";
  private static final String CATALOG_FILE_PATH = "catalog/catalog.xml";
  
  public static final Logger log = LogManager.getLogger(TestUseCatalogs.class.getSimpleName());

  
  @Test
  public void testCatalogResolution() throws Exception {

    ClassLoader classLoader = getClass().getClassLoader();
    File inFile = new File(classLoader.getResource("html/sample_web_page.html").getFile());
    File templateFile = new File(classLoader.getResource(DOTX_TEMPLATE_PATH).getFile());
    File xformFile = new File(classLoader.getResource("xsl/test-catalog-resolution.xsl").getFile());
    // NOTE: The running test uses the catalog as copied to the target/test-classes/catalog directory,
    //       not the source diretory, so if you're testing the catalog, e.g., in Oxygen, use 
    //       the copy in the target directory or the relative path won't resolve.
    File catalogFile = new File(classLoader.getResource(CATALOG_FILE_PATH).getFile());
    File outFile = new File("out/testCatalogResolution.docx");
    if (outFile.exists()) {
      outFile.delete();
    }

    boolean GOOD_OPTIONS = false;
    Options options = null;
    try {
      options = MakeDocx.buildOptions();
      GOOD_OPTIONS = true;
    } catch (Exception e) {
      // 
    }
    
    assertTrue("Options not good", GOOD_OPTIONS);
    
    assertTrue("No catalog option", options.hasLongOption("catalog"));
    assertTrue("No -k option", options.hasOption("k"));
    
    String[] args = { 
        "-k " + catalogFile.getAbsolutePath(), 
        "-i " + inFile.getAbsolutePath(), 
        "-o " + outFile.getAbsolutePath(), 
        "-t " + templateFile.getAbsolutePath(),
        "-x " + xformFile.getAbsolutePath() 

        };
    
    CommandLineParser parser = new DefaultParser();
    CommandLine cmd = parser.parse( options, args);
    
    String catalog = cmd.getOptionValue("k");
    assertNotNull("No catalog option value", catalog);
    assertEquals(catalogFile.getAbsolutePath(), catalog.trim());
    
    try {
      MakeDocx.handleCommandLine(options, args, log);
    } catch (Throwable e) {
      fail("Got exception from handleCommandLine(): " + e.getMessage());
    }

  }

}
