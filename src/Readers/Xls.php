<?php

namespace Chen\Excel\Readers;

use Chen\Excel\ExcelReader;
use Chen\Excel\OLEReader;

/**
 * Class for parsing XLS files
 *
 * @author Chenmobuys
 */
class Xls extends BaseReader
{
    /**
     * Load handle from filename
     *
     * @return void
     */
    public function load(string $filename, array $options = [])
    {
        $this->handle = new ExcelReader($filename, false, 'UTF-8');
        
        if (function_exists('mb_convert_encoding')) {
            $this->handle->setUTFEncoder('mb');
        }

        if (empty($this->sheetNames)) {
            foreach ($this->handle->boundsheets as $sheetIndex => $sheetInfo) {
                $this->sheetNames[$sheetIndex] = $sheetInfo['name'];
            }
        }

        $this->setSheetIndex($this->sheetIndex);
    }

    /**
     * Checks file is readable
     */
    public function isReadable(string $filename)
    {
        return (bool) (new OLEReader)->read($filename);
    }

    /**
    * Set sheet index by sheet index
    *
    * @param int $sheetIndex
    *
    * @return void
    */
    public function setSheetIndex(int $sheetIndex)
    {
        parent::setSheetIndex($sheetIndex);

        $this->columnCount = $this->handle->sheets[$this->sheetIndex]['numCols'];
        $this->rowCount = $this->handle->sheets[$this->sheetIndex]['numRows'];

        if (!$this->rowCount && count($this->handle->sheets[$this->sheetIndex]['cells'])) {
            end($this->handle->sheets[$this->sheetIndex]['cells']);
            $this->rowCount = (int) key($this->handle->sheets[$this->sheetIndex]['cells']);
        }

        if ($this->columnCount) {
            $this->emptyRow = array_fill(1, $this->columnCount, '');
        } else {
            $this->emptyRow = [];
        }
    }

    /**
     * Move forward to next element
     *
     * @return mixed
     */
    public function next()
    {
        //  Internal counter is advanced here instead of the if statement
        //	because apparently it's fully possible that an empty row will not be
        //	present at all
        $this->currentRowIndex++;

        if (isset($this->handle->sheets[$this->sheetIndex]['cells'][$this->currentRowIndex])) {
            $this->currentRow = $this->handle->sheets[$this->sheetIndex]['cells'][$this->currentRowIndex];

            if (!$this->currentRow) {
                return [];
            }

            $this->currentRow = $this->currentRow + $this->emptyRow;
            ksort($this->currentRow);

            $this->currentRow = array_values($this->currentRow);
            return $this->currentRow;
        } else {
            $this->currentRow = $this->emptyRow;
            return $this->currentRow;
        }
    }
}
