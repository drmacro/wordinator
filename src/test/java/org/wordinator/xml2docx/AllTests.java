package org.wordinator.xml2docx;

import org.junit.runner.RunWith;
import org.junit.runners.Suite;

/**
 * Root test suite for the project.
 *
 */
@RunWith(Suite.class)
@Suite.SuiteClasses({
	TestDocxGenerator.class,
	TestMeasurement.class,
	TestUseCatalogs.class
})

public final class AllTests {

}
