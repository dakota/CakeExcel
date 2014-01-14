<?php

App::uses('Controller', 'Controller');
App::uses('ExcelView', 'CakeExcel.View');

/**
 * Dummy controller
 */
class ExcelTestController extends Controller {
}

class ExcelViewTest extends CakeTestCase {

/**
 * setup callback
 *
 * @return void
 */
	public function setUp() {
		parent::setUp();
		$path = CakePlugin::path('CakeExcel') . 'Test' . DS . 'test_app' . DS . 'View' . DS;
		App::build(array('View' => $path, 'Vendor' => APP . DS . 'vendor' . DS));

		$Controller = new ExcelTestController();
		$this->View = new ExcelView($Controller);
	}

	public function tearDown() {
		unset($this->View);
		parent::tearDown();
	}

/**
 * testRender
 *
 */
	public function testConstruct() {
		$result = $this->View->response->type();
		$this->assertEquals('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', $result);

		$result = $this->View->options;
		$this->assertEquals(array('format' => 'Excel2007'), $result);

		$this->assertTrue($this->View->PhpExcel instanceof PhpExcel);
	}

	public function testRender() {
		$excelOutput = $this->View->render('index');

		//Valid Excel 2007 file
		$this->assertSame('504b030414', bin2hex(substr($excelOutput, 0, 5)));

		$result = $this->View->PhpExcel->getActiveSheet()->getCellByColumnAndRow(0, 0)->getValue();
		$this->assertEquals('Test string', $result);
	}

	public function testGetFilename() {
		$this->View->request->params['pass'][0] = 'Test';
		$result = $this->View->getFilename();
		$this->assertEquals('test.xlsx', $result);

		$this->View->options['filename'] = 'filename test';
		$result = $this->View->getFilename();
		$this->assertEquals('filename test.xlsx', $result);
	}

	public function testGetExtension() {
		$result = $this->View->getExtension();
		$this->assertEquals('xlsx', $result);
	}
}