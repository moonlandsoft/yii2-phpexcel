<?php

namespace moonland\phpexcel;

use Closure;

class TableThCell extends TableTdCell {


    /**
     * TableThCell constructor.
     * @param string|Closure $value
     * @param array $class
     * @param int $colspan
     * @param int $rowspan
     */
    public function __construct($value = '', array $class = [], int $colspan = 1, int $rowspan = 1)
    {
        parent::__construct($value, $class, $colspan, $rowspan);
        $this->tag = 'th';
    }


}