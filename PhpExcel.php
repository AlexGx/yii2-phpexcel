<?php

namespace alexgx\phpexcel;

/**
 * Class PhpExcel
 */
class PhpExcel extends \yii\base\Component
{
    protected $instance = null;

    public function getInstance()
    {
        if ($this->instance !== null) {
            $this->instance = new \PHPExcel();
        }
        return $this->instance;
    }
}
