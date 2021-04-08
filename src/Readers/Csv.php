<?php

namespace Chen\Excel\Readers;

/**
 * Class for parsing CSV files
 *
 * @author Chenmobuys
 */
class Csv extends BaseReader
{
    protected $options = [
        'delimiter' => '',
        'enclosure' => '"',
    ];

    private $encoding = 'UTF-8';

    private $bomLength = 0;

    /**
     * Load handle from filename
     *
     * @return void
     */
    public function load(string $filename, array $options = [])
    {
        $this->handle = fopen($filename, 'r');

        $this->sheetNames = [basename($filename)];

        $this->initBomLengthAndEncoding();

        $this->initDelimiter();
    }

    /**
     * Checks file is readable
     */
    public function isReadable(string $filename)
    {
        // Attempt to guess mimetype
        $type = mime_content_type($filename);
        
        $supportedTypes = [
            'application/csv',
            'text/csv',
            'text/plain',
            'inode/x-empty',
        ];

        return in_array($type, $supportedTypes, true);
    }

    /**
     * Init bomLength and Encoding
     */
    private function initBomLengthAndEncoding()
    {
        if (!$this->bomLength) {
            fseek($this->handle, 0);
            $BOM32 = bin2hex(fread($this->handle, 4));
            if ($BOM32 == '0000feff') {
                $this->encoding = 'UTF-32';
                $this->bomLength = 4;
            } elseif ($BOM32 == 'fffe0000') {
                $this->encoding = 'UTF-32';
                $this->bomLength = 4;
            }
        }

        fseek($this->handle, 0);
        $BOM8 = bin2hex(fread($this->handle, 3));
        if ($BOM8 == 'efbbbf') {
            $this->encoding = 'UTF-8';
            $this->bomLength = 3;
        }

        // Seeking the place right after BOM as the start of the real content
        if ($this->bomLength) {
            fseek($this->handle, $this->bomLength);
        }
    }

    /**
     * init delimiter
     */
    private function initDelimiter()
    {
        if (!$this->option('delimiter')) {
            // fgetcsv needs single-byte separators
            $semicolon = ';';
            $tab = "\t";
            $comma = ',';

            // Reading the first row and checking if a specific separator character
            // has more columns than others (it means that most likely that is the delimiter).
            $semicolonCount = count(fgetcsv($this->handle, null, $semicolon));
            fseek($this->handle, $this->bomLength);
            $tabCount = count(fgetcsv($this->handle, null, $tab));
            fseek($this->handle, $this->bomLength);
            $commaCount = count(fgetcsv($this->handle, null, $comma));
            fseek($this->handle, $this->bomLength);

            $delimiter = $semicolon;
            if ($tabCount > $semicolonCount || $commaCount > $semicolonCount) {
                $delimiter = $commaCount > $tabCount ? $comma : $tab;
            }

            $this->options['delimiter'] = $delimiter;
        }
    }

    /**
     * Rewind the Iterator to the first element
     *
     * @return void
     */
    public function rewind()
    {
        fseek($this->handle, $this->bomLength);
        parent::rewind();
    }

    /**
     * Move forward to next element
     *
     * @return mixed
     */
    public function next()
    {
        $this->currentRow = [];

        // Finding the place the next line starts for UTF-16 encoded files
        // Line breaks could be 0x0D 0x00 0x0A 0x00 and PHP could split lines on the
        //	first or the second linebreak leaving unnecessary \0 characters that mess up
        //	the output.
        if ($this->encoding == 'UTF-16LE' || $this->encoding == 'UTF-16BE') {
            while (!feof($this->handle)) {
                // While bytes are insignificant whitespace, do nothing
                $char = ord(fgetc($this->handle));
                if (!$char || $char == 10 || $char == 13) {
                    continue;
                } else {
                    // When significant bytes are found, step back to the last place before them
                    if ($this->encoding == 'UTF-16LE') {
                        fseek($this->handle, ftell($this->handle) - 1);
                    } else {
                        fseek($this->handle, ftell($this->handle) - 2);
                    }
                    break;
                }
            }
        }

        $this->currentRow = fgetcsv($this->handle, null, $this->option('delimiter'), $this->option('enclosure')) ?: [];

        if ($this->currentRow) {
            // Converting multi-byte unicode strings
            // and trimming enclosure symbols off of them because those aren't recognized
            // in the relevan encodings.
            if ($this->encoding != 'ASCII' && $this->encoding != 'UTF-8') {
                foreach ($this->currentRow as $key => $value) {
                    $this->currentRow[$key] = trim(trim(
                        mb_convert_encoding($value, 'UTF-8', $this->encoding),
                        $this->option('enclosure')
                    ));
                }
            }

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
        return $this->currentRow || !feof($this->handle);
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
}
