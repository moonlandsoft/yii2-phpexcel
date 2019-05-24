<?php

namespace moonland\phpexcel;

class TableThCell extends TableTdCell {

    public function __construct(string $value = '', array $class = [], int $colspan = 1, int $rowspan = 1)
    {
        parent::__construct($value, $class, $colspan, $rowspan);
        $this->tag = 'th';
    }


}