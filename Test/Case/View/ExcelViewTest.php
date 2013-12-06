<?php

App::uses('Controller', 'Controller');
App::uses('ExcelView', 'CakeExcel.View');

/**
 * Dummy controller
 */
class ExcelTestUsersController extends Controller {

	public $name = 'Users';
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
		App::build(array('View' => $path));

		$Controller = new ExcelTestUsersController();
		$this->View = new ExcelView($Controller);
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
}