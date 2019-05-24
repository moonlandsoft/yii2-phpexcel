<?php
namespace moonland\phpexcel;

class TableNavigator
{
    public $x;
    public $y;

    public $minX;
    public $minY;
    public $maxX;
    public $maxY;

    public $usedCells = [];

    public function __construct(int $x = 1, int $y = 1)
    {
        $this->x = $this->y = 1;
        $this->minX = $this->maxX = $x;
        $this->minY = $this->maxY = $y;
    }

    public function nextCell(): void
    {
        do {
            $this->x++;
        } while (isset($this->usedCells[$this->x][$this->y]));
        $this->maxX = max($this->x, $this->maxX);
    }

    public function newLine(): void
    {
        $this->x = $this->minX - 1;
        $this->y++;
        $this->nextCell();
        $this->maxY = max($this->y, $this->maxY);
    }

    public function setColspan(int $columns): void
    {
        $i = 1;
        while ($i < $columns) {
            $this->usedCells[$this->x + $i][$this->y] = [$this->x, $this->y];
            $i++;
        }
    }

    public function setRowspan(int $rows): void
    {
        $i = 1;
        while ($i < $rows) {
            $this->usedCells[$this->x][$this->y + $i] = [$this->x, $this->y];
            $i++;
        }
    }


}