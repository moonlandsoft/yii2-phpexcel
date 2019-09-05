<?php

namespace moonland\phpexcel;


use yii\helpers\Html;

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

    /** @var string */
    public $tooltipText;

    public $style = [];

    public $params;

    /**
     * TableCell constructor.
     * @param string $value *
     * @param array $class
     * @param $colspan
     * @param $rowspan
     */
    public function __construct(string $value = '', array $class = [], $colspan = 1, $rowspan = 1)
    {
        $this->value = $value;
        $this->colspan = $colspan;
        $this->rowspan = $rowspan;
        $this->class = $class;
    }

    public function setParam(string $name, $value): void
    {
        $this->params[$name] = $value;
    }


    public function getParam(string $name): bool
    {
        return $this->params[$name] ?? false;
    }

    public function getValue(): string
    {
        if ($this->numberDecimals !== false) {
            return number_format(round($this->value, $this->numberDecimals), $this->numberDecimals, '.', '');
        }
        return $this->value;
    }

    public function getHtml(): string
    {
        $options = [];
        if ($this->colspan !== 1) {
            $options['colspan'] = $this->colspan;
        }
        if ($this->rowspan !== 1) {
            $options['rowspan'] = $this->rowspan;
        }
        if ($this->class) {
            $options['class'] = $this->class;
        }
        if ($this->nowrap) {
            $options['class'] = array_merge($this->class, ['text-nowrap']);
        }

        if ($this->tooltipText) {
            $options['data-title'] = $this->tooltipText;
            $options['data-placement'] = 'top';
            $options['data-toggle'] = 'tooltip';
            $options['data-container'] = 'body';
            $options['data-original-title'] = '';
        }

        if(!$this->url) {
            return Html::tag($this->tag, $this->getValue(), $options);
        }

        return Html::tag($this->tag, Html::a($this->getValue(), $this->url, $this->urlOptions), $options);
    }


}