<?php

namespace Chen\Excel\Readers;

use Iterator;
use Countable;
use SimpleXMLElement;
use XMLReader;
use Chen\Excel\ExcelReader;
use Chen\Excel\ReaderException;

/**
 * BaseReader
 *
 * @author Chenmbouys
 */
abstract class BaseReader implements Iterator, Countable
{
    /**
     * @var ExcelReader|XMLReader|SimpleXMLElement $handle
     */
    protected $handle;

    /**
     * @var array $options
     */
    protected $options;

    /**
     * @var int $rowCount
     */
    protected $rowCount = 0;

    /**
     * @var int $columnCount
     */
    protected $columnCount = 0;

    /**
     * @var array $sheetNames
     */
    protected $sheetNames = [];

    /**
     * @var int $sheetIndex
     */
    protected $sheetIndex = 0;

    /**
     * @var array $currentRow
     */
    protected $currentRow = [];

    /**
     * @var int $currentRowIndex
     */
    protected $currentRowIndex = 0;

    /**
     * @var array $emptyRow
     */
    protected $emptyRow = [];

    /**
     * Load handle from filename
     *
     * @param string $filename
     * @param array $options
     * 
     * @return void
     */
    abstract public function load($filename, $options = []);

    /**
     * Checks file is readable
     * 
     * @param string $filename
     * 
     * @return bool
     */
    abstract public function isReadable($filename);

    /**
     * Get option value by option name
     *
     * @return mixed
     */
    public function option($name, $default = null)
    {
        return isset($this->options[$name]) ? $this->options[$name] : $default;
    }

    /**
     * Get Sheet Names
     *
     * @return array
     */
    public function getSheetNames()
    {
        return $this->sheetNames;
    }

    /**
     * Get Sheet Indexes
     *
     * @return mixed
     */
    public function getSheetIndexes()
    {
        return array_keys($this->sheetNames);
    }

    /**
    * Set sheet index by sheet index
    *
    * @param int $sheetIndex
    *
    * @return void
    */
    public function setSheetIndex($sheetIndex)
    {
        if (!in_array($sheetIndex, array_keys($this->sheetNames))) {
            throw new ReaderException('Sheet index:' . $sheetIndex .' does not exist');
        }
        $this->sheetIndex = $sheetIndex;
        $this->rewind();
    }

    /**
     * Set sheet index by sheet name
     *
     * @param string $sheetName
     *
     * @return void
     */
    public function setSheetIndexByName($sheetName)
    {
        $sheetNamesReverse = array_reverse($this->sheetNames);

        if (!isset($sheetNamesReverse[$sheetName])) {
            throw new ReaderException('Sheet name:' . $sheetName .' does not exist');
        }

        $this->setSheetIndex($sheetNamesReverse[$sheetName]);
    }

    /**
     * Rewind the Iterator to the first element
     *
     * @return void
     */
    public function rewind()
    {
        $this->currentRow = [];
        $this->currentRowIndex = 0;
    }

    /**
     * Return the current element.
     * Similar to the current() function for arrays in PHP
     *
     * @return mixed
     */
    public function current()
    {
        if ($this->currentRowIndex == 0 && empty($this->currentRow)) {
            $this->next();
        }
        return $this->currentRow;
    }

    /**
     * Move forward to next element
     *
     * @return mixed
     */
    public function next()
    {
        $this->currentRowIndex++;

        return $this->currentRow;
    }
    
    /**
     * Return the key of the current element
     *
     * @return int
     */
    public function key()
    {
        return $this->currentRowIndex;
    }

    /**
     * Checks if current position is valid
     *
     * @return bool
     */
    public function valid()
    {
        return $this->currentRowIndex <= $this->rowCount;
    }

    /**
     * Count elements of an object
     *
     * @return int
     */
    public function count()
    {
        return $this->rowCount;
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        unset($this->handle);
    }
}
