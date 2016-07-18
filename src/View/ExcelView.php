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
     * Excel layouts are located in the xlsx sub directory of `Layouts/`
     *
     * @var string
     */
    public $layoutPath = 'xlsx';

    /**
     * Excel views are always located in the 'xlsx' sub directory for a
     * controllers views.
     *
     * @var string
     */
    public $subDir = 'xlsx';

    /**
     * PHPExcel instance
     * @var PhpExcel
     */
    public $PhpExcel = null;

    /**
     * Constructor
     *
     * @param \Cake\Network\Request $request Request instance.
     * @param \Cake\Network\Response $response Response instance.
     * @param \Cake\Event\EventManager $eventManager EventManager instance.
     * @param array $viewOptions An array of view options
     */
    public function __construct(
        Request $request = null,
        Response $response = null,
        EventManager $eventManager = null,
        array $viewOptions = []
    ) {
        if (!empty($viewOptions['templatePath']) && $viewOptions['templatePath'] == '/xlsx') {
            $this->subDir = null;
        }

        parent::__construct($request, $response, $eventManager, $viewOptions);

        if (isset($viewOptions['name']) && $viewOptions['name'] == 'Error') {
            $this->subDir = null;
            $this->layoutPath = null;
            $response->type('html');

            return;
        }

        if ($response && $response instanceof Response) {
            $response->type('xlsx');
        }

        $this->PhpExcel = new PHPExcel();
    }

    /**
     * Render method
     *
     * @param string $view The view being rendered.
     * @param string $layout The layout being rendered.
     * @return string The rendered view.
     */
    public function render($view = null, $layout = null)
    {
        $content = parent::render($view, $layout);
        if ($this->response->type() == 'text/html') {
            return $content;
        }

        $this->Blocks->set('content', $this->output());
        $this->response->download($this->getFilename());

        return $this->Blocks->get('content');
    }

    /**
     * Generates the binary excel data
     *
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
     * Gets the filename
     *
     * @return string filename
     */
    public function getFilename()
    {
        if (isset($this->viewVars['_filename'])) {
            return $this->viewVars['_filename'] . '.xlsx';
        }

        return Inflector::slug(str_replace('.xlsx', '', $this->request->url)) . '.xlsx';
    }
}
