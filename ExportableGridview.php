<?php

namespace eseperio\gridview;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Yii;
use yii\base\UserException;
use yii\grid\Column;
use yii\helpers\ArrayHelper;
use yii\helpers\Html;
use yii\helpers\Url;

/**
 * Class ExportableGridview
 * @package eseperio\admintheme\widgets\grid
 */
class ExportableGridview extends \yii\grid\GridView
{

    const DEFAULT_POINTER = 'A';

    const WRITER_XLS = 'Xls';
    const WRITER_XLSX = 'Xlsx';
    const WRITER_ODS = 'Ods';
    const WRITER_CSV = 'Csv';
    const WRITER_HTML = 'Html';
    const WRITER_TCPDF = 'Tcpdf';
    const WRITER_DOMPDF = 'Dompdf';
    const WRITER_MPDF = 'Mpdf';
    
    /**
     * Maximum iterations for cleaning output buffers to prevent infinite loops
     */
    const MAX_BUFFER_CLEANUP_ITERATIONS = 100;
    /**
     * @var string the layout that determines how different sections of the grid view should be organized.
     * The following tokens will be replaced with the corresponding section contents:
     *
     * - `{summary}`: the summary section. See [[renderSummary()]].
     * - `{errors}`: the filter model error summary. See [[renderErrors()]].
     * - `{items}`: the list items. See [[renderItems()]].
     * - `{sorter}`: the sorter. See [[renderSorter()]].
     * - `{pager}`: the pager. See [[renderPager()]].
     */
    public $layout = "{summary}\n{items}\n{export}\n{pager}";
    /**
     * @var string writer type (format type). If not set, it will be determined automatically.
     * Supported values:
     *
     * - 'Xls'
     * - 'Xlsx'
     * - 'Ods'
     * - 'Csv'
     * - 'Html'
     * - 'Tcpdf'
     * - 'Dompdf'
     * - 'Mpdf'
     *
     * @see IOFactory
     */
    public $writerType;
    /**
     * @var string filename of the generated spreadsheet
     */
    public $fileName = 'exported.xls';
    /**
     * @var array Additional options for sending the file
     */
    public $exportFileOptions = [];
    /**
     * @var bool whether the gridview can be exported.
     */
    public $exportable = true;
    /**
     * @var array options to use when rendering export link.
     */
    public $exportLinkOptions = [
        'class' => 'btn btn-default',
        'target' => '_blank',
    ];
    /**
     * @var array columns to be exported. It empty gridview columns will be used.
     */
    public $exportColumns = [];
    /**
     * @var string spreadsheet column index
     */
    private $columnIndex = self::DEFAULT_POINTER;
    /**
     * @var int spreadsheet row index
     */
    private $rowIndex = 1;
    /**
     * @var array multidimensional containing rows and columns.
     *            First level: Rows
     *            Second level: Cols
     */
    private $data = [];

    /**
     * @var Spreadsheet generated
     */
    private $_document;

    /**
     * Initialize the gridview;
     * @throws \yii\base\InvalidConfigException
     */
    public function init()
    {

        if (!isset($this->id)) {
            $this->options['id'] = $this->getId();
        }

        if ($this->downloadRequested() && !empty($this->exportColumns)) {
            $this->emptyCell = "";
            $this->columns = $this->exportColumns;
        }
        if ($this->downloadRequested())
            $this->dataProvider->pagination = false;

        parent::init();
    }

    /**
     * @return bool whether a download is allowed and requested.
     */
    public function downloadRequested()
    {
        $request = Yii::$app->getRequest();

        $grid = $request->get('export-grid');
        $container = $request->get('export-container');

        return $this->exportable && $grid && $container === $this->id;
    }

    /**
     * @return string|void
     * @throws UserException
     * @throws \Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \yii\base\ExitException
     */
    public function run()
    {

        if ($this->downloadRequested()) {
            if ($this->dataProvider->getCount() <= 0 || empty($this->columns))
                throw new UserException('Nothing to export');

            $response = Yii::$app->getResponse();
            
            // Clean output buffers to prevent HTML from being included in the export,
            // but preserve the outermost buffer level to maintain compatibility with
            // testing frameworks like Codeception. We remove all nested buffers (level 2+)
            // and then clean the content of the outermost buffer (level 1), while keeping
            // the buffer structure intact for the testing framework.
            $iterations = 0;
            while (ob_get_level() > 1 && $iterations++ < self::MAX_BUFFER_CLEANUP_ITERATIONS) {
                $result = ob_end_clean();
                if ($result === false) {
                    // If ob_end_clean fails, break to avoid infinite loop
                    break;
                }
            }
            if (ob_get_level() > 0) {
                ob_clean();
            }
            
            $response->setStatusCode(200);
            
            // Prepare the export data
            $this->prepareExportArray();
            $document = $this->getDocument();
            $document->getActiveSheet()->fromArray($this->data);
            
            // Prepare the response for file sending
            $this->prepareSend($this->exportFileOptions);
            
            // Send the response and end the application
            Yii::$app->response->send();
            Yii::$app->end();

        }

        parent::run();
    }

    protected function prepareExportArray()
    {
        $this->renderExportHeaders();
        $this->renderExportBody();
        $this->renderExportFooter();
        $this->cleanExportData();
    }

    public function renderExportHeaders()
    {
        $cells = [];
        foreach ($this->columns as $column) {
            /* @var $column Column */
            $cells[$this->columnIndex++] = $column->renderHeaderCell();
        }
        $this->data[$this->rowIndex++] = $cells;
        $this->columnIndex = self::DEFAULT_POINTER;
    }

    public function renderExportBody()
    {
        $models = array_values($this->dataProvider->getModels());
        $keys = $this->dataProvider->getKeys();
        $rows = [];
        foreach ($models as $index => $model) {
            $key = $keys[$index];
            $rows[] = $this->renderExportRow($model, $key, $index);
        }

    }

    public function renderExportRow($model, $key, $index)
    {
        $cells = [];
        /* @var $column Column */
        foreach ($this->columns as $column) {
            $cells[$this->columnIndex++] = $column->renderDataCell($model, $key, $index);
        }
        $this->columnIndex = self::DEFAULT_POINTER;

        $this->data[$this->rowIndex++] = $cells;
    }

    public function renderExportFooter()
    {
        $cells = [];

        foreach ($this->columns as $column) {
            /* @var $column Column */
            $cells[$this->columnIndex++] = $column->renderFooterCell();
        }
        $this->data[$this->rowIndex++] = $cells;
        $this->columnIndex = self::DEFAULT_POINTER;


    }

    /**
     * Removes all tags and encodes each cell to export.
     */
    private function cleanExportData()
    {
        foreach ($this->data as $rowKey => $row) {
            foreach ($row as $colKey => $column) {
                $cleanValue = Html::encode(strip_tags(trim(str_replace('&nbsp;', "", $column))));
                $this->data[$rowKey][$colKey] = $cleanValue;
            }
        }
    }

    /**
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet spreadsheet document representation instance.
     */
    public function getDocument()
    {
        if (!is_object($this->_document)) {
            $this->_document = new Spreadsheet();
        }

        return $this->_document;
    }

    /**
     * Sends the rendered content as a file to the browser.
     *
     * Note that this method only prepares the response for file sending. The file is not sent
     * until [[\yii\web\Response::send()]] is called explicitly or implicitly.
     * The latter is done after you return from a controller action.
     *
     * @param array $options additional options for sending the file. The following options are supported:
     *
     *  - `mimeType`: the MIME type of the content. Defaults to 'application/octet-stream'.
     *  - `inline`: bool, whether the browser should open the file within the browser window. Defaults to false,
     *    meaning a download dialog will pop up.
     *
     * @return \yii\web\Response the response object.
     */
    public function prepareSend($options = [])
    {

        $writerType = $this->writerType;
        if ($writerType === null) {
            $fileExtension = strtolower(pathinfo($this->fileName, PATHINFO_EXTENSION));
            $writerType = ucfirst($fileExtension);
        }

        $tmpResource = tmpfile();
        if ($tmpResource === false)
            throw new \Exception('Temporary file could not be created');

        $tmpResourceMetaData = stream_get_meta_data($tmpResource);
        $tmpFileName = $tmpResourceMetaData['uri'];

        $writer = IOFactory::createWriter($this->getDocument(), $writerType);
        $writer->save($tmpFileName);
        unset($writer);

        return Yii::$app->getResponse()->sendStreamAsFile($tmpResource, $this->fileName, $options);
    }

    /**
     * @inheritdoc
     */
    public function renderSection($name)
    {

        if ($name === '{export}' && $this->exportable)
            return $this->renderExportLink();

        return parent::renderSection($name);
    }

    /**
     * @return string the export link tag.
     */
    public function renderExportLink()
    {
        $label = ArrayHelper::remove($this->exportLinkOptions, 'label', 'Export');
        $encode = ArrayHelper::remove($this->exportLinkOptions, 'encode', true);
        $url = Url::current(['export-grid' => 1, 'export-container' => $this->getId()]);
        $this->exportLinkOptions['data-pjax'] = 0;

        return Html::a($encode ? Html::encode($label) : $label, $url, $this->exportLinkOptions);
    }


}
