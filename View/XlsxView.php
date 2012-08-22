<?php
App::uses('View', 'View');

class XlsxView extends View {
	protected $PhpExcel = null;
	private $options = array(
		'graphTemplate' => false,
		'format' => 'xlsx'
	);

/**
 * Constructor
 *
 * @param Controller $controller
 * @return void
 */
	public function __construct(Controller $Controller = null) {
		parent::__construct($Controller);

		$this->response->type($this->request->params['ext']);
		if ($Controller instanceof CakeErrorController) {
			$this->response->type('html');
		}

		if($this->request->params['ext'] == 'xls') {
			$this->options['format'] = 'xls';
		}

		//Excel files can get big!
		ini_set('memory_limit', '2048M');
		set_time_limit(0);

		App::import('Vendor', 'CakeExcel.PhpExcel', array('file' => 'PHPExcel' . DS . 'Classes' . DS . 'PHPExcel' . DS . 'IOFactory.php'));

		$this->PhpExcel = new PHPExcel();
	}

	public function setOption($key, $value) {
		$this->options[$key] = $value;
	}

	public function render($action = null, $layout = null, $file = null) {
		$this->viewPath .= DS . 'xlsx';

		parent::render($action, $layout, $file);

		if($this->options['graphTemplate'] && $this->options['format'] == 'xlsx') {
			$this->graphOutput();
		}
		else {
			$response = $this->standardOutput();
		}

		$id = current($this->request->params['pass']);
		$filename = strtolower($this->viewPath) . $id;
		if (isset($this->options['filename'])) {
			$filename = $this->options['filename'];
		}

		$this->response->download($filename . '.' . $this->request->params['ext']);

		$this->Blocks->set('content', $response);
		return $this->Blocks->get('content');
	}

	private function standardOutput() {
		ob_start();
		
		if($this->options['format'] == 'xlsx') {
			$objWriter = PHPExcel_IOFactory::createWriter($this->PhpExcel, 'Excel2007');
			$extension = 'xlsx';
		}
		elseif($this->options['format'] == 'xls') {
			$objWriter = PHPExcel_IOFactory::createWriter($this->PhpExcel, 'Excel5');
			$extension = 'xls';
		}

		if(isset($objWriter)) {
			$objWriter->setPreCalculateFormulas(false);
			$objWriter->save('php://output');
		}

		return ob_get_clean();
	}

	private function graphOutput() {
		try {
			$tmpLocation = '/tmp/' . time() . rand(0, time()) . '/';

			$graphFile = APP . $this->options['graphTemplate'];

			if(is_readable($graphFile)) {
				$this->unzip($graphFile, $tmpLocation . 'template/');
			}
			else {
				throw new Exception('Could not read template.');
			}

			$objWriter = new PHPExcel_Writer_Excel2007($this->PhpExcel);
			$objWriter->save($tmpLocation . 'source.xlsx');
			$unzippedSource = $this->unzip($tmpLocation . 'source.xlsx', $tmpLocation . 'source/');

			$this->copyFiles($tmpLocation);

			if($this->zip($tmpLocation . 'final.xlsx', $tmpLocation . 'template/')) {
				$file = $tmpLocation . 'final.xlsx';
				header('Content-Description: File Transfer');
				header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
				header('Content-Disposition: attachment;filename="'.$this->options['filename'].'.xlsx"');
				header('Content-Transfer-Encoding: binary');
				header('Cache-Control: max-age=0');
				header('Content-Length: ' . filesize($file));
				ob_clean();
				flush();
				readfile($file);
			}
		}
		catch (Exception $e) {
			die($e->getMessage());
		}

		$this->cleanUp($tmpLocation);
	}

	private function cleanUp($dir) {
		$files = scandir($dir);
		array_shift($files);		// remove '.' from array
		array_shift($files);		// remove '..' from array

		foreach($files as $file) {
			$file = $dir . '/' . $file;
			if(is_dir($file)) {
				$this->cleanUp($file);
				rmdir($file);
			}
			else {
				unlink($file);
			}
		}
		rmdir($dir);
	}

	private function listdir($dir='.') {
		if(!is_dir($dir)) {
			return false;
		}

		$dir = rtrim($dir, '/');
		$files = array();
		$this->listdiraux($dir, $files);

		return $files;
	}

	private function listdiraux($dir, &$files) {
		$handle = opendir($dir);
		while(($file = readdir($handle)) !== false) {
			if($file == '.' || $file == '..') {
				continue;
			}
			$filepath = $dir == '.' ? $file : $dir . '/' . $file;

			if(is_link($filepath)){
				continue;
			}

			if(is_file($filepath)) {
				$files[] = $filepath;
			}
			else if(is_dir($filepath)) {
				$this->listdiraux($filepath, $files);
			}
		}
		closedir($handle);
	}

	private function zip ($filename, $location) {
		$files = $this->listdir($location);
		sort($files, SORT_LOCALE_STRING);

		$zip = new ZipArchive();
		if($zip->open($filename, ZIPARCHIVE::OVERWRITE) !== true) {
			throw new Exception("Could not created zip file $filename");
		}
		//add the files
		foreach($files as $file) {
			$zip->addFile($file, str_replace($location, '', $file));
		}

		$zip->close();

		return file_exists($filename);
	}

	private function copyFiles($location) {
		$dest = $location . 'template/xl/worksheets/sheet1.xml';
		$source = $location . 'source/xl/worksheets/sheet1.xml';

		if(!copy($source, $dest)){
			throw new Exception("failed to copy $file...");
		}

		$dest = $location . 'template/xl/sharedStrings.xml';
		$source = $location . 'source/xl/sharedStrings.xml';

		if(!copy($source, $dest)){
			throw new Exception("failed to copy $file...");
		}
	}

	private function unzip($fileName, $location) {
		$zip = new ZipArchive;

		$res = $zip->open($fileName);
		if($res === true) {
			$zip->extractTo($location);
			$zip->close();
			return;
		}
		else{
			throw new Exception('Could not unzip template. Error code: ' . $res);
		}
	}
}