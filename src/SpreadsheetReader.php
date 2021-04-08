<?php

namespace Chen\Excel;

use Chen\Excel\Readers\BaseReader;
use InvalidArgumentException;

/**
 * Main class for spreadsheet reading
 *
 * @author Chenmobuys
 */
class SpreadsheetReader
{
    private static $readers = [
        'Csv' => Readers\Csv::class,
        'Ods' => Readers\Ods::class,
        'Xls' => Readers\Xls::class,
        'Xlsx' => Readers\Xlsx::class,
        // TODO next to
        // 'Xml' => Readers\Xml::class,
        // 'Slk' => Reader\Slk::class,
        // 'Gnumeric' => Reader\Gnumeric::class,
        // 'Html' => Reader\Html::class,
    ];

    /**
     * Create Reader For ReaderType
     *
     * @param string $readerType
     *
     * @return BaseReader
     */
    public static function createReader($readerType)
    {
        if (!isset(self::$readers[$readerType])) {
            throw new InvalidArgumentException("No reader found for type $readerType");
        }

        // Instantiate reader
        $className = self::$readers[$readerType];

        return new $className();
    }

    /**
     * Create Reader For File
     *
     * @param string $filename
     *
     * @return BaseReader
     */
    public static function createReaderForFile($filename)
    {
        if (!is_file($filename)) {
            throw new InvalidArgumentException('File "' . $filename . '" does not exist.');
        }

        if (!is_readable($filename)) {
            throw new InvalidArgumentException('Could not open "' . $filename . '" for reading.');
        }

        $guessedReader = self::getReaderTypeFromExtension($filename);

        if ($guessedReader !== null) {
            $reader = self::createReader($guessedReader);

            // Let's see if we are lucky
            if ($reader->isReadable($filename)) {
                return $reader;
            }
        }

        foreach (self::$readers as $type => $class) {
            //    Ignore our original guess, we know that won't work
            if ($type !== $guessedReader) {
                $reader = self::createReader($type);
                if ($reader->isReadable($filename)) {
                    return $reader;
                }
            }
        }

        throw new ReaderException('Unable to identify a reader for this file');
    }

    /**
     * Load file reader
     *
     * @param string $filename
     * @param array $options
     * 
     * @return BaseReader
     */
    public static function load($filename, $options = [])
    {
        $reader = self::createReaderForFile($filename);
        $reader->load($filename, $options);
        return $reader;
    }

    /**
     * Get ReaderType from Extension
     *
     * @param $filename
     *
     * @return string $readerType
     *
     */
    private static function getReaderTypeFromExtension($filename)
    {
        $extension = strtolower(pathinfo($filename, PATHINFO_EXTENSION));

        if (is_null($extension)) {
            return null;
        }

        switch (strtolower($extension)) {
            case 'xlsx': // Excel (OfficeOpenXML) Spreadsheet
            case 'xlsm': // Excel (OfficeOpenXML) Macro Spreadsheet (macros will be discarded)
            case 'xltx': // Excel (OfficeOpenXML) Template
            case 'xltm': // Excel (OfficeOpenXML) Macro Template (macros will be discarded)
                return 'Xlsx';
            case 'xls': // Excel (BIFF) Spreadsheet
            case 'xlt': // Excel (BIFF) Template
                return 'Xls';
            case 'ods': // Open/Libre Offic Calc
            case 'ots': // Open/Libre Offic Calc Template
                return 'Ods';
            case 'slk':
                return 'Slk';
            case 'xml': // Excel 2003 SpreadSheetML
                return 'Xml';
            case 'gnumeric':
                return 'Gnumeric';
            case 'htm':
            case 'html':
                return 'Html';
            case 'csv':
            case 'tsv':
                return 'Csv';
            default:
                return null;
        }
    }

    /**
    * Register a reader with its type and class name.
    *
    * @param string $readerType
    * @param string $readerClass
    */
    public static function registerReader($readerType, $readerClass)
    {
        if (!is_a($readerClass, BaseReader::class, true)) {
            throw new InvalidArgumentException('Registered readers must implement ' . BaseReader::class);
        }

        self::$readers[$readerType] = $readerClass;
    }
}
