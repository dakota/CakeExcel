<?php
namespace CakeExcel\View;

use Cake\Event\EventManager;
use Cake\Http\Response;
use Cake\Http\ServerRequest;
use Cake\Utility\Text;
use Cake\View\View;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

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
     * Spreadsheet instance
     *
     * @var Spreadsheet
     */
    public $Spreadsheet = null;

    /**
     * Constructor
     *
     * @param \Cake\Network\Request $request Request instance.
     * @param \Cake\Network\Response $response Response instance.
     * @param \Cake\Event\EventManager $eventManager EventManager instance.
     * @param array $viewOptions An array of view options
     */
    public function __construct(
        ServerRequest $request = null,
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
            $this->response = $this->response->withType('html');

            return;
        }

        if ($response && $response instanceof Response) {
            $this->response = $this->response->withType('xlsx');
        }

        $this->Spreadsheet = new Spreadsheet();
    }

    /**
     * Magic accessor for helpers. Backward compatibility for PHPExcel property
     *
     * @param string $name Name of the attribute to get.
     *
     * @return mixed
     */
    public function __get($name)
    {
        if ($name === 'PhpExcel') {
            return $this->Spreadsheet;
        }

        return parent::__get($name);
    }

    /**
     * Render method
     *
     * @param string|false|null $view Name of view file to use
     * @param string|null $layout Layout to use.
     * @return string|null Rendered content or null if content already rendered and returned earlier.
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function render($view = null, $layout = null)
    {
        $content = parent::render($view, $layout);
        if ($this->response->getType() == 'text/html') {
            return $content;
        }

        $this->Blocks->set('content', $this->output());
        $this->response = $this->response->withDownload($this->getFilename());

        return $this->Blocks->get('content');
    }

    /**
     * Generates the binary excel data
     *
     * @return string
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    protected function output()
    {
        ob_start();

        $writer = new Xlsx($this->Spreadsheet);

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

        return Text::slug(str_replace('.xlsx', '', $this->request->getRequestTarget())) . '.xlsx';
    }
}
