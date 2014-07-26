<?php

namespace alexgx\phpexcel;

/**
 * Class PhpExcel
 */
class PhpExcel extends \yii\base\Object
{
    /**
     * @var string
     */
    public $defaultFormat = 'Excel2007';

    /**
     * Creates new workbook
     * @return \PHPExcel
     */
    public function create()
    {
        return new \PHPExcel();
    }

    /**
     * @param string $filename name of the spreadsheet file
     * @return \PHPExcel
     */
    public function load($filename)
    {
        return \PHPExcel_IOFactory::load($filename);
    }

    /**
     *
     * NOTE: enchasement, resolve name on file extension
     * @param \PHPExcel $object
     * @param string $name attachment name
     * @param string $format output format
     */
    public function responseFile(\PHPExcel $object, $filename, $format = null)
    {
        if ($format === null) {
            $format = $this->resolveFormat($filename);
        }
        $writer = \PHPExcel_IOFactory::createWriter($object, $format);
        ob_start();
        $writer->save('php://output');
        $content = ob_get_clean();
        \Yii::$app->response->sendContentAsFile($content, $filename, $this->resolveMime($format));
        \Yii::$app->end();
    }

    public function writeData($config)
    {
        $object = $this->create();
        $sheet = $object->getActiveSheet();
        $config['sheet'] = &$sheet;
        $writer = new ExcelDataWriter($config);
        $writer->write();
        return $object;
    }

    /**
     *
     * @param $format
     * @return string
     */
    protected function resolveMime($format)
    {
        // TODO: add additional types (formats)
        $list = [
            'OpenDocument' => 'application/vnd.oasis.opendocument.spreadsheet',
            'Excel5' => 'application/vnd.ms-excel',
            'Excel2007' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        ];
        return isset($list[$format]) ? $list[$format] : 'application/octet-stream';
    }

    /**
     *
     * @param $filename
     * @return string
     */
    protected function resolveFormat($filename)
    {
        $list = [
            'ods' => 'OpenDocument',
            'xls' => 'Excel5',
            'xlsx' => 'Excel2007',
        ];
        // TODO: check strtolower
        $extension = pathinfo($filename, PATHINFO_EXTENSION);
        return isset($list[$extension]) ? $list[$extension] : $this->defaultFormat;
    }
}
