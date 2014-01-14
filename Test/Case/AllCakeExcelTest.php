<?php
/**
 * All CakeExcel plugin tests
 */
class AllCakeExcelTest extends CakeTestCase {

/**
 * Suite define the tests for this plugin
 *
 * @return void
 */
	public static function suite() {
		$suite = new CakeTestSuite('All CakeExcel test');

		$path = CakePlugin::path('CakeExcel') . 'Test' . DS . 'Case' . DS;
		$suite->addTestDirectoryRecursive($path);

		return $suite;
	}

}
