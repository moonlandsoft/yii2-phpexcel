<?php

namespace moonland\phpexcel;


class TableTdCell
{
    public $tag = 'td';
    private $value;
    public $colspan;
    public $rowspan;
    public $class = [];
    public $nowrap = false;
    public $url = [];
    public $urlOptions = [];
    public $numberDecimals = false;
    public $columnAutoSize = true;

    public $style = [];

    public $params;

    /**
     * TableCell constructor.
     * @param string $value*
     * @param array $class
     * @param $colspan
     * @param $rowspan

     */
    public function __construct(string $value = '', array $class =[], $colspan = 1,  $rowspan = 1)
    {
        $this->value = $value;
        $this->colspan = $colspan;
        $this->rowspan = $rowspan;
        $this->class = $class;
    }

    public function setParam(string $name, $value)
    {
        $this->params[$name] = $value;
    }


    public function getParam(string $name)
    {
        return $this->params[$name]??false;
    }

    public function getValue()
    {
        if($this->numberDecimals !== false) {
            return number_format(round($this->value, $this->numberDecimals), $this->numberDecimals,'.','');
        }
        return $this->value;
    }

}