<?php

namespace alexgx\phpexcel;

use yii\helpers\ArrayHelper;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

class ExcelDataWriter extends \yii\base\BaseObject
{
    /**
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public $sheet;

    /**
     * @var array
     */
    public $data = [];

    /**
     * @var array
     */
    public $columns = [];

    /**
     * @var array
     */
    public $options = [];

    /**
     * @var string
     */
    public $defaultDateFormat = \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_DDMMYYYY;

    /**
     * @var int Start row Default value is `1`
     */
    protected $j = 1;

    /**
     * @var int
     */
    protected $startRow = 1;

    /**
     * @var int start column
     */
    protected  $startColumn = 1;

    /**
     * @var int
     */
    protected  $endColumn = 1;

    /**
     * @var bool
     */
    protected $freezeHeader = false;

    /**
     * @var bool
     */
    protected $autoFilter = false;



    /**
     * @param bool $freezeHeader
     */
    public function setFreezeHeader($freezeHeader)
    {
        $this->freezeHeader = $freezeHeader;
    }



    /**
     * @param int $startRow
     */
    public function setStartRow($startRow){
        $this->j = $startRow;
        $this->startRow = $startRow;
    }

    /**
     * @param int $startColumn
     */
    public function setStartColumn($startColumn){
        $this->startColumn = $startColumn;
    }



    public function write()
    {
        if (!is_array($this->data) || !is_array($this->columns)) {
            return;
        }

        $this->writeHeaderRow();
        $this->writeDataRows();
        $this->writeFooterRow();
    }

    protected function writeHeaderRow()
    {

        if($this->freezeHeader){
            $startColumnAsString = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($this->startColumn);
            $this->sheet->freezePane($startColumnAsString . ($this->j+1));
        }
        $i = $this->startColumn;
        foreach ($this->columns as $column) {
            $this->endColumn = $i + 1;
            if (isset($column['header'])) {
                $this->sheet->setCellValueByColumnAndRow($i, $this->j, $column['header']);
            }
            if (isset($column['headerStyles'])) {
                $this->sheet->getStyleByColumnAndRow($i, $this->j)->applyFromArray($column['headerStyles']);
            }

            if (isset($column['width']) && $column['width'] !== 'autosize') {
                $this->sheet->getColumnDimensionByColumn($i)->setWidth($column['width']);
            }
            ++$i;
        }
        ++$this->j;
    }

    protected function writeDataRows()
    {
        foreach ($this->data as $key => $row) {
            $i = $this->startColumn;
            if (isset($this->options['rowOptions']) && $this->options['rowOptions'] instanceof \Closure) {
                $rowOptions = call_user_func($this->options['rowOptions'], $row, $key);
            }
            foreach ($this->columns as $column) {
                if (isset($rowOptions)) {
                    $column = ArrayHelper::merge($column, $rowOptions);
                }
                if (isset($column['cellOptions']) && $column['cellOptions'] instanceof \Closure) {
                    $column = ArrayHelper::merge($column, call_user_func($column['cellOptions'], $row, $key, $i, $this->j));
                }
                $value = null;
                if (isset($column['value'])) {
                    $value = ($column['value'] instanceof \Closure) ? call_user_func($column['value'], $row, $key) : $column['value'];
                } elseif (isset($column['attribute']) && isset($row[$column['attribute']])) {
                    $value = $row[$column['attribute']];
                }

                $this->writeCell($value, $i, $this->j, $column);
                ++$i;
            }

            ++$this->j;
        }
    }

    protected function writeFooterRow()
    {
        $i = $this->startColumn;
        foreach ($this->columns as $column) {
            if (isset($column['width']) && $column['width'] === 'autosize') {
                $this->sheet->getColumnDimensionByColumn($this->j)->setAutoSize(true);
            }
            // footer config
            $config = [];
            if (isset($column['footerStyles'])) {
                $config['styles'] = $column['footerStyles'];
            }
            if (isset($column['footerType'])) {
                $config['type'] = $column['footerType'];
            }
            if (isset($column['footerLabel'])) {
                $config['label'] = $column['footerLabel'];
            }
            if (isset($column['footerOptions']) && $column['footerOptions'] instanceof \Closure) {
                $config = ArrayHelper::merge($config, call_user_func($column['footerOptions'], null, null, $i, $this->j));
            }
            $value = null;
            if (isset($column['footer'])) {
                $value = ($column['footer'] instanceof \Closure) ? call_user_func($column['footer'], null, null) : $column['footer'];
            }
            $this->writeCell($value, $i, $this->j, $config);


            ++$i;
        }

        ++$this->j;
    }

    protected function writeCell($value, $column, $row, $config)
    {
        // auto type
        if (!isset($config['type']) || $config['type'] === null) {
            $this->sheet->setCellValueByColumnAndRow($column, $row, $value);
        } elseif ($config['type'] === 'date') {
            $timestamp = !is_int($value) ? strtotime($value) : $value;
            $this->sheet
                ->getStyleByColumnAndRow($column, $row)
                ->getNumberFormat()
                ->setFormatCode(NumberFormat::FORMAT_DATE_DATETIME);


            $this->sheet->setCellValueByColumnAndRow($column, $row, \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel($timestamp));
            if (!isset($config['styles']['numberformat']['code'])) {
                $config['styles']['numberformat']['code'] = $this->defaultDateFormat;
            }

            $this->sheet
                ->getStyleByColumnAndRow($column, $row)
                ->getNumberFormat()
                ->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);

        } elseif ($config['type'] === 'url') {
            if (isset($config['label'])) {
                if ($config['label'] instanceof \Closure) {
                    // NOTE: calculate label on top level
                    $label = call_user_func($config['label']/*, TODO */);
                } else {
                    $label = $config['label'];
                }
            } else {
                $label = $value;
            }
            $urlValid = (filter_var($value, FILTER_VALIDATE_URL) !== false);
            if (!$urlValid) {
                $label = '';
            }

            $this->sheet->setCellValueByColumnAndRow($column, $row, $label);

            if ($urlValid) {
                $this->sheet->getCellByColumnAndRow($column, $row)->getHyperlink()->setUrl($value);
            }
        } else {
            $this->sheet->setCellValueExplicitByColumnAndRow($column, $row, $value, $config['type']);
        }

        if (isset($config['styles'])) {
            $this->sheet->getStyleByColumnAndRow($column, $row)->applyFromArray($config['styles']);
        }

        if (isset($config['format']['decimal'])) {
            if($config['format']['decimal'] === 0){
                $numberFormat = NumberFormat::FORMAT_NUMBER;
            }else{
                $numberFormat = '0.' . str_repeat('0',$config['format']['decimal']);
            }
            $this->sheet
                ->getStyleByColumnAndRow($column, $row)
                ->getNumberFormat()
                ->setFormatCode($numberFormat);
        }

    }

    /**
     * @return int
     */
    public function getJ()
    {
        return $this->j;
    }
}
