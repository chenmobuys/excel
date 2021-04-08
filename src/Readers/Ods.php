<?php

namespace Chen\Excel\Readers;

use XMLReader;
use ZipArchive;
use InvalidArgumentException;

/**
 * Class for parsing ODS files
 *
 * @author Chenmobuys
 */
class Ods extends BaseReader
{
    /**
     * @var array $options
     */
    protected $options = [
        'tempDir' => '',
        'returnDateTimeObjects' => false,
    ];

    /**
     * @var string $tempDir
     */
    private $tempDir;

    /**
     * @var string $contentPath
     */
    private $contentPath;

    /**
     * @var bool $tableOpen
     */
    private $tableOpen = false;

    /**
     * @var bool $rowOpen
     */
    private $rowOpen = false;

    /**
     * @var bool $valid
     */
    private $valid = false;

    /**
     * Load handle from filename
     *
     * @param string $filename
     * @param array $options
     * 
     * @return void
     */
    public function load($filename, $options = [])
    {
        $this->options = array_merge($this->options, ['tempDir' => sys_get_temp_dir()], $options);

        $this->tempDir = $this->option('tempDir');
        $this->tempDir = rtrim($this->tempDir, DIRECTORY_SEPARATOR);
        $this->tempDir = $this->tempDir . DIRECTORY_SEPARATOR . uniqid() . DIRECTORY_SEPARATOR;

        $zip = new ZipArchive;

        if (!($status = $zip->open($filename))) {
            throw new InvalidArgumentException('SpreadsheetReader ODS: File not readable ('.$filename.') (Error '.$status.')');
        }

        if ($zip->locateName('content.xml') !== false) {
            $zip->extractTo($this->tempDir, 'content.xml');
            $this->contentPath = $this->tempDir . 'content.xml';
        }

        $zip->close();

        if ($this->contentPath && is_readable($this->contentPath)) {
            if (empty($this->sheetNames)) {
                $this->handle = new XMLReader;
                $this->handle->open($this->contentPath);
    
                while ($this->handle->read()) {
                    if ($this->handle->name == 'table:table') {
                        $this->sheetNames[] = $this->handle->getAttribute('table:name');
                        $this->handle->next();
                    }
                }
                $this->handle->close();
            }

            $this->handle = new XMLReader;
            $this->handle->open($this->contentPath);
            $this->valid = true;
            $this->setSheetIndex($this->sheetIndex);
        }
    }

    /**
     * Checks file is readable
     * 
     * @param string $filename
     * 
     * @return bool
     */
    public function isReadable($filename)
    {
        $zip = new ZipArchive();
        $result = (bool) $zip->open($filename);
        if ($result) {
            $zip->close();
        }
        
        return $result;
    }

    /**
     * Rewind the Iterator to the first element
     *
     * @return void
     */
    public function rewind()
    {
        if ($this->currentRowIndex > 0) {
            // If the worksheet was already iterated, XML file is reopened.
            // Otherwise it should be at the beginning anyway
            $this->handle->close();
            $this->handle->open($this->contentPath);
            $this->valid = true;
            $this->tableOpen = false;
            $this->rowOpen = false;

            parent::rewind();
        }

        $this->currentRowIndex = 0;
    }

    /**
     * Move forward to next element
     *
     * @return mixed
     */
    public function next()
    {
        $this->currentRow = [];

        if (!$this->tableOpen) {
            $tableCounter = 0;
            $skipRead = false;

            while ($this->valid = ($skipRead || $this->handle->read())) {
                if ($skipRead) {
                    $skipRead = false;
                }

                if ($this->handle->name == 'table:table' && $this->handle->nodeType != XMLReader::END_ELEMENT) {
                    if ($tableCounter == $this->sheetIndex) {
                        $this->tableOpen = true;
                        break;
                    }

                    $tableCounter++;
                    $this->handle->next();
                    $skipRead = true;
                }
            }
        }

        if ($this->tableOpen && !$this->rowOpen) {
            while ($this->valid = $this->handle->read()) {
                switch ($this->handle->name) {
                    case 'table:table':
                        $this->tableOpen = false;
                        $this->handle->next('office:document-content');
                        $this->valid = false;
                        break 2;
                    case 'table:table-row':
                        if ($this->handle->nodeType != XMLReader::END_ELEMENT) {
                            $this->rowOpen = true;
                            break 2;
                        }
                        break;
                }
            }
        }
        

        if ($this->rowOpen) {
            $lastCellContent = '';

            while ($this->valid = $this->handle->read()) {
                switch ($this->handle->name) {
                    case 'table:table-cell':
                        if ($this->handle->nodeType == XMLReader::END_ELEMENT || $this->handle->isEmptyElement) {
                            if ($this->handle->nodeType == XMLReader::END_ELEMENT) {
                                $cellValue = $lastCellContent;
                            } elseif ($this->handle->isEmptyElement) {
                                $lastCellContent = '';
                                $cellValue = $lastCellContent;
                            }

                            $this->currentRow[] = $cellValue;

                            if ($this->handle->getAttribute('table:number-columns-repeated') !== null) {
                                $repeatedColumnCount = $this->handle->getAttribute('table:number-columns-repeated');
                                // Checking if larger than one because the value is already added to the row once before
                                if ($repeatedColumnCount > 1) {
                                    $this->currentRow = array_pad($this->currentRow, count($this->currentRow) + $repeatedColumnCount - 1, $lastCellContent);
                                }
                            }
                        } else {
                            $lastCellContent = '';
                        }
                        // no break
                    case 'text:p':
                        if ($this->handle->nodeType != XMLReader::END_ELEMENT) {
                            $lastCellContent = $this->handle->readString();
                        }
                        break;
                    case 'table:table-row':
                        $this->rowOpen = false;
                        break 2;
                }
            }
        }

        if ($this->currentRow) {
            return parent::next();
        }

        return $this->currentRow;
    }

    /**
     * Checks if current position is valid
     *
     * @return bool
     */
    public function valid()
    {
        return $this->valid;
    }

    /**
    * Count elements of an object
    *
    * @return int
    */
    public function count()
    {
        return $this->currentRowIndex;
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        if ($this->handle && $this->handle instanceof XMLReader) {
            $this->handle->close();
        }
        
        @unlink($this->contentPath);

        parent::__destruct();
    }
}
