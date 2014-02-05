<?php
/**
 * *
 * Licensed under The MIT License
 * For full copyright and license information, please see the LICENSE.txt
 * Redistributions of files must retain the above copyright notice.
 *
 * @license       http://www.opensource.org/licenses/mit-license.php MIT License
 */

App::uses('View', 'View');

/**
 * @package  Cake.View
 */
class ExcelView extends View {

/**
 * PHPExcel instance
 * @var PhpExcel
 */
	public $PhpExcel = null;

	protected $_extensions = array(
		'Excel5' => 'xls',
		'Excel2007' => 'xlsx'
	);

/**
 * Options for the renderer
 * @var array
 */
	public $options = array(
		'format' => 'Excel2007'
	);

/**
 * Constructor
 *
 * @param Controller $controller
 * @return void
 */
	public function __construct(Controller $Controller = null) {
		parent::__construct($Controller);

		if (isset($this->request->params['ext']) && $this->request->params['ext'] == 'xls') {
			$this->options['format'] = 'Excel5';
		}

		$this->__setResponseType();

		if ($Controller instanceof CakeErrorController) {
			$this->response->type('html');
		}

		//PhpExcel instances can get big!
		ini_set('memory_limit', '2048M');
		set_time_limit(0);

		if (!class_exists('PHPExcel')) {
			App::import('Vendor', 'PhpExcel', array('file' => 'phpoffice' . DS . 'phpexcel' . DS . 'Classes' . DS . 'PHPExcel' . DS . 'IOFactory.php'));
		}

		$this->PhpExcel = new PHPExcel();
	}

/**
 * [render description]
 * @param  [type] $action [description]
 * @param  [type] $layout [description]
 * @param  [type] $file   [description]
 * @return [type]         [description]
 */
	public function render($action = null, $layout = null, $file = null) {
		$this->viewPath .= DS . 'Excel';

		$content = parent::render($action, false, $file);
		if ($this->response->type() == 'text/html') {
			return $content;
		}
		$content = $this->__output();
		$this->Blocks->set('content', $content);

		$this->response->download($this->getFilename());

		return $this->Blocks->get('content');
	}

/**
 * Generates the binary excel data
 * @return string
 * @throws CakeException If the excel writer does not exist
 */
	private function __output() {
		ob_start();

		$writer = PHPExcel_IOFactory::createWriter($this->PhpExcel, $this->options['format']);

		if (!isset($writer)) {
			throw new CakeException(__d('cake_excel', 'Excel writer "%s" not found'), $this->options['format']);
		}

		$writer->setPreCalculateFormulas(false);
		$writer->setIncludeCharts(true);
		$writer->save('php://output');

		$output = ob_get_clean();

		return $output;
	}

	public function getFilename() {
		if (!empty($this->options['filename'])) {
			return $this->options['filename'] . '.' . $this->getExtension();
		}
		$id = current($this->request->params['pass']);
		return strtolower($id) . '.' . $this->getExtension();
	}

	public function getExtension() {
		return $this->_extensions[$this->options['format']];
	}

	private function __setResponseType() {
		$this->response->type($this->_extensions[$this->options['format']]);
	}
}