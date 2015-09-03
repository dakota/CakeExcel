<?php
namespace CakeExcel\View;

use Cake\Core\Exception\Exception;
use Cake\Event\EventManager;
use Cake\Network\Request;
use Cake\Network\Response;
use Cake\Utility\Inflector;
use Cake\View\View;
use PHPExcel;
use PHPExcel_IOFactory;

/**
 * @package  Cake.View
 */
class ExcelView extends View
{

    /**
     * PHPExcel instance
     * @var PhpExcel
     */
    public $PhpExcel = null;

    /**
     * Filename String
     * @var string
     */
    protected $filename;

    /**
     * SubDir Name
     * @var string
     */
    public $subDir = 'xlsx';

    /**
     * Constructor
     *
     * @param \Cake\Network\Request $request Request instance.
     * @param \Cake\Network\Response $response Response instance.
     * @param \Cake\Event\EventManager $eventManager Event manager instance.
     * @param array $viewOptions View options. See View::$_passedVars for list of
     *   options which get set as class properties.
     *
     * @throws \Cake\Core\Exception\Exception
     */
    public function __construct(Request $request = null, Response $response = null, EventManager $eventManager = null, array $viewOptions = [])
    {
        parent::__construct($request, $response, $eventManager, $viewOptions);

        if (isset($viewOptions['name']) && $viewOptions['name'] == 'Error') {
            $this->subDir = null;
            $this->layoutPath = null;
            $response->type('html');
            return;
        }

        $this->PhpExcel = new PHPExcel();
    }

    /**
     * Render method
     * @param  string $action - action to render
     * @param  string $layout - layout to use
     * @return string - rendered content
     */
    public function render($view = null, $layout = null)
    {
        $content = parent::render($view, $layout);
        if ($this->response->type() == 'text/html') {
            return $content;
        }

        $content = $this->output();
        $this->Blocks->set('content', $content);
        $this->response->download($this->getFilename());
        return $this->Blocks->get('content');
    }

    /**
     * Generates the binary excel data
     * @return string
     * @throws CakeException If the excel writer does not exist
     */
    protected function output()
    {
        ob_start();

        $writer = PHPExcel_IOFactory::createWriter($this->PhpExcel, 'Excel2007');

        if (!isset($writer)) {
            throw new Exception('Excel writer not found');
        }

        $writer->setPreCalculateFormulas(false);
        $writer->setIncludeCharts(true);
        $writer->save('php://output');

        $output = ob_get_clean();

        return $output;
    }

    /**
     * Sets the filename
     * @param string $filename the filename
     * @return void
     */
    public function setFilename($filename)
    {
        $this->filename = $filename;
    }

    /**
     * Gets the filename
     * @return string filename
     */
    public function getFilename()
    {
        if (!empty($this->filename)) {
            return $this->filename . '.xlsx';
        }
        return Inflector::slug(str_replace('.xlsx', '', $this->request->url)) . '.xlsx';
    }
}
