<?php

namespace alexgx\phpexcel;

/**
 * Class PhpExcel
 */
class PhpExcel extends \yii\base\BaseObject
{
    /**
     * @var string
     */
    public $defaultFormat = 'Excel2007';

    /**
     * Creates new workbook
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    public function create()
    {
        return new \PhpOffice\PhpSpreadsheet\Spreadsheet();
    }

	/**
	 * Creates new Worksheet Drawing
	 * @return \PhpOffice\PhpSpreadsheet\Worksheet\Drawing
	 */
	public function getObjDrawing() {
		return new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
	}

    /**
     * @param string $filename name of the spreadsheet file
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    public function load($filename)
    {
        return \PhpOffice\PhpSpreadsheet\IOFactory::load($filename);
    }

    /**
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet $object
     * @param string $name attachment name
     * @param string $format output format
     */
    public function responseFile(\PhpOffice\PhpSpreadsheet\Spreadsheet $object, $filename, $format = null, $beautifyFileName = true)
    {
        if ($format === null) {
            $format = $this->resolveFormat($filename);
        }

        $filename = $this->filterFilename($filename, $beautifyFileName);
        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($object, $format);
        ob_start();
        $writer->save('php://output');
        $content = ob_get_clean();
        \Yii::$app->response->sendContentAsFile($content, $filename, $this->resolveMime($format));
        \Yii::$app->end();
    }

    /**
     * @param $sheet
     * @param $config
     */
    public function writeSheetData($sheet, $data, $config)
    {
        $config['sheet'] = &$sheet;
        $config['data'] = $data;
        $writer = new ExcelDataWriter($config);
        $writer->write();
        return $sheet;
    }

    public function writeTemplateData(/* TODO */)
    {
        // TODO: implement
    }

    public function readSheetData($sheet, $config)
    {
        // TODO: implement
    }

    /**
     *
     * @param $format
     * @return string
     */
    protected function resolveMime($format)
    {
        $list = [
            'CSV' => 'text/csv',
            'HTML' => 'text/html',
            'PDF' => 'application/pdf',
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
        // see IOFactory::createReaderForFile etc.
        $list = [
            'ods' => 'Ods',
            'xls' => 'Xls',
            'xlsx' => 'Xlsx',
            'csv' => 'Csv',
            'pdf' => 'Pdf',
            'html' => 'Html',
        ];
        // TODO: check strtolower
        $extension = pathinfo($filename, PATHINFO_EXTENSION);
        return isset($list[$extension]) ? $list[$extension] : $this->defaultFormat;
    }

    /**
     * sanitise file name
     *
     * @param string $filename
     * @param bool $beautify
     * @return string
     */
    public function filterFilename($filename, $beautify = true)
    {
        // sanitize filename
        $filename = preg_replace(
            '~
        [<>:"/\\|?*]|            # file system reserved https://en.wikipedia.org/wiki/Filename#Reserved_characters_and_words
        [\x00-\x1F]|             # control characters http://msdn.microsoft.com/en-us/library/windows/desktop/aa365247%28v=vs.85%29.aspx
        [\x7F\xA0\xAD]|          # non-printing characters DEL, NO-BREAK SPACE, SOFT HYPHEN
        [#\[\]@!$&\'()+,;=]|     # URI reserved https://tools.ietf.org/html/rfc3986#section-2.2
        [{}^\~`]                 # URL unsafe characters https://www.ietf.org/rfc/rfc1738.txt
        ~x',
            '-', $filename);
        // avoids ".", ".." or ".hiddenFiles"
        $filename = ltrim($filename, '.-');
        // optional beautification
        if ($beautify) $filename = $this->beautifyFilename($filename);
        // maximise filename length to 255 bytes http://serverfault.com/a/9548/44086
        $ext = pathinfo($filename, PATHINFO_EXTENSION);
        $filename = mb_strcut(pathinfo($filename, PATHINFO_FILENAME), 0, 255 - ($ext ? strlen($ext) + 1 : 0), mb_detect_encoding($filename)) . ($ext ? '.' . $ext : '');
        return $filename;
    }

    /**
     *
     * @param string $filename
     * @return string
     */
    function beautifyFilename($filename)
    {
        // reduce consecutive characters
        $filename = preg_replace(array(
            // "file   name.zip" becomes "file-name.zip"
            '/ +/',
            // "file___name.zip" becomes "file-name.zip"
            '/_+/',
            // "file---name.zip" becomes "file-name.zip"
            '/-+/'
        ), '-', $filename);
        $filename = preg_replace(array(
            // "file--.--.-.--name.zip" becomes "file.name.zip"
            '/-*\.-*/',
            // "file...name..zip" becomes "file.name.zip"
            '/\.{2,}/'
        ), '.', $filename);
        // lowercase for windows/unix interoperability http://support.microsoft.com/kb/100625
        $filename = mb_strtolower($filename, mb_detect_encoding($filename));
        // ".file-name.-" becomes "file-name"
        $filename = trim($filename, '.-');
        return $filename;
    }
}
