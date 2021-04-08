<?php

namespace Chen\Excel\Readers;

use ZipArchive;
use XMLReader;
use SimpleXMLElement;
use DateTime;
use DateTimeZone;
use DateInterval;
use InvalidArgumentException;

/**
 * Class for parsing XLSX files specifically
 *
 * @author Chenmobuys
 */
class Xlsx extends BaseReader
{
    const CELL_TYPE_BOOL = 'b';
    const CELL_TYPE_NUMBER = 'n';
    const CELL_TYPE_ERROR = 'e';
    const CELL_TYPE_SHARED_STR = 's';
    const CELL_TYPE_STR = 'str';
    const CELL_TYPE_INLINE_STR = 'inlineStr';

    protected $options = [
        'tempDir' => '',
        'returnDateTimeObjects' => false,
        'sharedStringsCacheLimit' => null,
    ];

    /**
     * @var string $tempDir
     */
    private $tempDir;

    /**
     * @var array $tempFiles
     */
    private $tempFiles = [];

    /**
     * @var XMLReader $sharedStrings
     */
    private $sharedStrings;

    /**
     * @var string $sharedStringsPath
     */
    private $sharedStringsPath;

    /**
     * @var string $SharedStringsIndex
     */
    private $sharedStringsIndex = 0;

    /**
     * @var int $sharedStringsCount
     */
    private $sharedStringsCount = 0;

    /**
     * @var bool $sharedStringsOpen
     */
    private $sharedStringsOpen = false;
    
    /**
     * @var bool $sharedStringsForwarded
     */
    private $sharedStringsForwarded = false;

    /**
     * @var string $lastSharedStringValue
     */
    private $lastSharedStringsValue;

    /**
     * @var array $sharedStringsCache
     */
    private $sharedStringsCache = [];

    /**
     * @var SimpleXMLElement $styles
     */
    private $styles;

    /**
     * @var array $stylesCache
     */
    private $stylesCache = [];

    /**
     * @var array $parsedFormatCache
     */
    private $parsedFormatCache = [];

    /**
     * @var array $formats
     */
    private $formats = [];

    /**
     * @var XMLReader $worksheet
     */
    private $worksheet;

    /**
     * @var bool $valid
     */
    private $valid;

    /**
     * @var bool $rowOpen
     */
    private $rowOpen;

    /**
     * @var bool $GMPSupported
     */
    private $GMPSupported;

    /**
     * @var DateTime $baseDate
     */
    private $baseDate;

    /**
     * @var string $decimalSeparator
     */
    private $decimalSeparator;

    /**
     * @var string $thousandSeparator
     */
    private $thousandSeparator;

    /**
     * @var string $currencyCode
     */
    private $currencyCode;

    /**
     * @var array $builtinFormats
     */
    private static $builtinFormats = [
        0 => '',
        1 => '0',
        2 => '0.00',
        3 => '#,##0',
        4 => '#,##0.00',

        9 => '0%',
        10 => '0.00%',
        11 => '0.00E+00',
        12 => '# ?/?',
        13 => '# ??/??',
        14 => 'mm-dd-yy',
        15 => 'd-mmm-yy',
        16 => 'd-mmm',
        17 => 'mmm-yy',
        18 => 'h:mm AM/PM',
        19 => 'h:mm:ss AM/PM',
        20 => 'h:mm',
        21 => 'h:mm:ss',
        22 => 'm/d/yy h:mm',

        37 => '#,##0 ;(#,##0)',
        38 => '#,##0 ;[Red](#,##0)',
        39 => '#,##0.00;(#,##0.00)',
        40 => '#,##0.00;[Red](#,##0.00)',

        45 => 'mm:ss',
        46 => '[h]:mm:ss',
        47 => 'mmss.0',
        48 => '##0.0E+0',
        49 => '@',

        // CHT & CHS
        27 => '[$-404]e/m/d',
        30 => 'm/d/yy',
        36 => '[$-404]e/m/d',
        50 => '[$-404]e/m/d',
        57 => '[$-404]e/m/d',

        // THA
        59 => 't0',
        60 => 't0.00',
        61 =>'t#,##0',
        62 => 't#,##0.00',
        67 => 't0%',
        68 => 't0.00%',
        69 => 't# ?/?',
        70 => 't# ??/??'
    ];
    
    /**
     * @var array $dateReplacements
     */
    private static $dateReplacements = [
        'All' => [
            '\\' => '',
            'am/pm' => 'A',
            'yyyy' => 'Y',
            'yy' => 'y',
            'mmmmm' => 'M',
            'mmmm' => 'F',
            'mmm' => 'M',
            ':mm' => ':i',
            'mm' => 'm',
            'm' => 'n',
            'dddd' => 'l',
            'ddd' => 'D',
            'dd' => 'd',
            'd' => 'j',
            'ss' => 's',
            '.s' => ''
        ],
        '24H' => [
            'hh' => 'H',
            'h' => 'G'
        ],
        '12H' => [
            'hh' => 'h',
            'h' => 'G'
        ]
    ];
   
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
            throw new InvalidArgumentException('SpreadsheetReader Xlsx: File not readable "'.$filename.'", Message: '. $status .'');
        }

        // Getting the general workbook information
        if ($zip->locateName('xl/workbook.xml') !== false) {
            $this->handle = new SimpleXMLElement($zip->getFromName('xl/workbook.xml'));
        }

        // Extracting the XMLs from the XLSX zip file
        if ($zip->locateName('xl/sharedStrings.xml') !== false) {
            $this->sharedStringsPath = $this->tempDir . 'xl'.DIRECTORY_SEPARATOR.'sharedStrings.xml';
            $zip->extractTo($this->tempDir, 'xl/sharedStrings.xml');
            $this->tempFiles[] = $this->tempDir .'xl'.DIRECTORY_SEPARATOR.'sharedStrings.xml';

            if (is_readable($this->sharedStringsPath)) {
                $this->sharedStrings = new XMLReader;
                $this->sharedStrings->open($this->sharedStringsPath);
                $this->prepareSharedStringCache();
            }
        }

        if (empty($this->sheetNames) && is_object($this->handle->sheets)) {
            foreach ($this->handle->sheets->sheet as $index => $sheet) {
                $attributes = $sheet->attributes('r', true);
                foreach ($attributes as $name => $value) {
                    if ($name == 'id') {
                        $sheetId = (int) str_replace('rId', '', (string) $value);
                        break;
                    }
                }
                $this->sheetNames[$sheetId] = (string) $sheet['name'];
            }
            ksort($this->sheetNames);

            foreach ($this->sheetNames as $sheetIndex => $sheetName) {
                $this->sheetIndex = $sheetIndex;
                break;
            }
        }

        foreach ($this->sheetNames as $index => $name) {
            $zip->extractTo($this->tempDir, 'xl/worksheets/sheet' . $index . '.xml');
            if ($zip->locateName('xl/worksheets/sheet'.$index.'.xml') !== false) {
                $zip->extractTo($this->tempDir, 'xl/worksheets/sheet' . $index . '.xml');
                $this->tempFiles[] = $this->tempDir . 'xl' . DIRECTORY_SEPARATOR . 'worksheets' . DIRECTORY_SEPARATOR . 'sheet' . $index . '.xml';
            }
        }

        // If worksheet is present and is OK, parse the styles already
        if ($zip->locateName('xl/styles.xml') !== false) {
            $this->styles = new SimpleXMLElement($zip->getFromName('xl/styles.xml'));
            if ($this->styles && $this->styles->cellXfs && $this->styles->cellXfs->xf) {
                foreach ($this->styles->cellXfs->xf as $index => $xf) {
                    // Format #0 is a special case - it is the "General" format that is applied regardless of applyNumberFormat
                    if ($xf->attributes()->applyNumberFormat || (0 == (int) $xf->attributes()->numFmtId)) {
                        $formatId = (int)$xf->attributes()->numFmtId;
                        // If format ID >= 164, it is a custom format and should be read from styleSheet\numFmts
                        $this->stylesCache[] = $formatId;
                    } else {
                        // 0 for "General" format
                        $this->stylesCache[] = 0;
                    }
                }
            }
            
            if ($this->styles->numFmts && $this->styles->numFmts->numFmt) {
                foreach ($this->styles->numFmts->numFmt as $Index => $numFmt) {
                    $this->formats[(int)$numFmt->attributes()->numFmtId] = (string) $numFmt->attributes()->formatCode;
                }
            }
            unset($this->styles);
        }

        $zip->close();

        $this->GMPSupported = function_exists('gmp_gcd');

        // Setting base date
        if (!$this->baseDate) {
            $this->baseDate = new DateTime;
            $this->baseDate->setTimezone(new DateTimeZone('UTC'));
            $this->baseDate->setDate(1900, 1, 0);
            $this->baseDate->setTime(0, 0, 0);
        }

        // Decimal and thousand separators
        if (!$this->decimalSeparator && !$this->thousandSeparator && !$this->currencyCode) {
            $locale = localeconv();
            $this->decimalSeparator = $locale['decimal_point'];
            $this->thousandSeparator = $locale['thousands_sep'];
            $this->currencyCode = $locale['int_curr_symbol'];
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
        $zip->close();
        
        return $result;
    }
    
    /**
     * Creating shared string cache if the number of shared strings is acceptably low (or there is no limit on the amount)
     *
     * @return void
     */
    private function prepareSharedStringCache()
    {
        while ($this->sharedStrings->read()) {
            if ($this->sharedStrings->name == 'sst') {
                $this->sharedStringsCount = $this->sharedStrings->getAttribute('count');
                break;
            }
        }

        if (!$this->sharedStringsCount
            || (
                $this->option('sharedStringsCacheLimit') < $this->sharedStringsCount
                && $this->option('sharedStringsCacheLimit') !== null
            )
        ) {
            return;
        }

        $cacheIndex = 0;
        $cacheValue = '';
        while ($this->sharedStrings->read()) {
            switch ($this->sharedStrings->name) {
                case 'si':
                    if ($this->sharedStrings->nodeType == XMLReader::END_ELEMENT) {
                        $this->sharedStringsCache[$cacheIndex] = $cacheValue;
                        $cacheIndex++;
                        $cacheValue = '';
                    }
                    break;
                case 't':
                    if ($this->sharedStrings->nodeType == XMLReader::END_ELEMENT) {
                        break;
                    }
                    $cacheValue .= $this->sharedStrings->readString();
                    break;
            }
        }

        $this->sharedStrings->close();
    }

    /**
     * Retrieves a shared string value by its index
     *
     * @param int $index
     *
     * @return string
     */
    private function getSharedString($index)
    {
        if (($this->option('sharedStringsCacheLimit') === null || $this->option('sharedStringsCacheLimit') > 0) && !empty($this->sharedStringsCache)) {
            if (isset($this->sharedStringsCache[$index])) {
                return $this->sharedStringsCache[$index];
            } else {
                return '';
            }
        }

        // If the desired index is before the current, rewind the XML
        if ($this->sharedStringsIndex > $index) {
            $this->sharedStringsOpen = false;
            $this->sharedStrings->close();
            $this->sharedStrings->open($this->sharedStringsPath);
            $this->sharedStringsIndex = 0;
            $this->lastSharedStringsValue = null;
            $this->sharedStringsForwarded = false;
        }

        // Finding the unique string count (if not already read)
        if ($this->sharedStringsIndex == 0 && !$this->sharedStringsCount) {
            while ($this->sharedStrings->read()) {
                if ($this->sharedStrings->name == 'sst') {
                    $this->sharedStringsCount = $this->sharedStrings->getAttribute('uniqueCount');
                    break;
                }
            }
        }

        // If index of the desired string is larger than possible, don't even bother.
        if ($this->sharedStringsCount && ($index >= $this->sharedStringsCount)) {
            return '';
        }

        // If an index with the same value as the last already fetched is requested
        // (any further traversing the tree would get us further away from the node)
        if (($index == $this->sharedStringsIndex) && ($this->lastSharedStringsValue !== null)) {
            return $this->lastSharedStringsValue;
        }

        // Find the correct <si> node with the desired index
        while ($this->sharedStringsIndex <= $index) {
            // SSForwarded is set further to avoid double reading in case nodes are skipped.
            if ($this->sharedStringsForwarded) {
                $this->sharedStringsForwarded = false;
            } else {
                $readStatus = $this->sharedStrings->read();
                if (!$readStatus) {
                    break;
                }
            }

            if ($this->sharedStrings->name == 'si') {
                if ($this->sharedStrings->nodeType == XMLReader::END_ELEMENT) {
                    $this->sharedStringsOpen = false;
                    $this->sharedStringsIndex++;
                } else {
                    $this->sharedStringsOpen = true;

                    if ($this->sharedStringsIndex < $index) {
                        $this->sharedStringsOpen = false;
                        $this->sharedStrings->next('si');
                        $this->sharedStringsForwarded = true;
                        $this->sharedStringsIndex++;
                    }
                    break;
                }
            }
        }

        $value = '';

        // Extract the value from the shared string
        if ($this->sharedStringsOpen && ($this->sharedStringsIndex == $index)) {
            while ($this->sharedStrings->read()) {
                switch ($this->sharedStrings->name) {
                    case 't':
                        if ($this->sharedStrings->nodeType == XMLReader::END_ELEMENT) {
                            break;
                        }
                        $value .= $this->sharedStrings->readString();
                        break;
                    case 'si':
                        if ($this->sharedStrings->nodeType == XMLReader::END_ELEMENT) {
                            $this->sharedStringsOpen = false;
                            $this->sharedStringsForwarded = true;
                            break 2;
                        }
                        break;
                }
            }
        }

        if ($value) {
            $this->lastSharedStringsValue = $value;
        }
        return $value;
    }

    /**
     * Formats the value according to the index
     *
     * @param string $value Cell value
     * @param int $index Format index
     *
     * @return string Formatted cell value
     */
    private function formatValue($value, $index)
    {
        if (!is_numeric($value)) {
            return $value;
        }

        if (isset($this->stylesCache[$index]) && ($this->stylesCache[$index] !== false)) {
            $index = $this->stylesCache[$index];
        } else {
            return $value;
        }

        // A special case for the "General" format
        if ($index == 0) {
            return $this->generalFormat($value);
        }

        $format = [];

        if (isset($this->parsedFormatCache[$index])) {
            $format = $this->parsedFormatCache[$index];
        }

        if (!$format) {
            $format = [
                'Code' => false,
                'Type' => false,
                'Scale' => 1,
                'Thousands' => false,
                'Currency' => false
            ];

            if (isset(self::$builtinFormats[$index])) {
                $format['Code'] = self::$builtinFormats[$index];
            } elseif (isset($this->formats[$index])) {
                $format['Code'] = $this->formats[$index];
            }

            // Format code found, now parsing the format
            if ($format['Code']) {
                $sections = explode(';', $format['Code']);
                $format['Code'] = $sections[0];

                switch (count($sections)) {
                    case 2:
                        if ($value < 0) {
                            $format['Code'] = $sections[1];
                        }
                        break;
                    case 3:
                    case 4:
                        if ($value < 0) {
                            $format['Code'] = $sections[1];
                        } elseif ($value == 0) {
                            $format['Code'] = $sections[2];
                        }
                        break;
                }
            }

            // Stripping colors
            $format['Code'] = trim(preg_replace('{^\[[[:alpha:]]+\]}i', '', $format['Code']));

            // Percentages
            if (substr($format['Code'], -1) == '%') {
                $format['Type'] = 'Percentage';
            } elseif (preg_match('{^(\[\$[[:alpha:]]*-[0-9A-F]*\])*[hmsdy]}i', $format['Code'])) {
                $format['Type'] = 'DateTime';

                $format['Code'] = trim(preg_replace('{^(\[\$[[:alpha:]]*-[0-9A-F]*\])}i', '', $format['Code']));
                $format['Code'] = strtolower($format['Code']);

                $format['Code'] = strtr($format['Code'], self::$dateReplacements['All']);
                if (strpos($format['Code'], 'A') === false) {
                    $format['Code'] = strtr($format['Code'], self::$dateReplacements['24H']);
                } else {
                    $format['Code'] = strtr($format['Code'], self::$dateReplacements['12H']);
                }
            } elseif ($format['Code'] == '[$EUR ]#,##0.00_-') {
                $format['Type'] = 'Euro';
            } else {
                // Removing skipped characters
                $format['Code'] = preg_replace('{_.}', '', $format['Code']);
                // Removing unnecessary escaping
                $format['Code'] = preg_replace("{\\\\}", '', $format['Code']);
                // Removing string quotes
                $format['Code'] = str_replace(['"', '*'], '', $format['Code']);
                // Removing thousands separator
                if (strpos($format['Code'], '0,0') !== false || strpos($format['Code'], '#,#') !== false) {
                    $format['Thousands'] = true;
                }
                $format['Code'] = str_replace(['0,0', '#,#'], ['00', '##'], $format['Code']);

                // Scaling (Commas indicate the power)
                $scale = 1;
                $matches = [];
                if (preg_match('{(0|#)(,+)}', $format['Code'], $matches)) {
                    $scale = pow(1000, strlen($matches[2]));
                    // Removing the commas
                    $format['Code'] = preg_replace(['{0,+}', '{#,+}'], ['0', '#'], $format['Code']);
                }

                $format['Scale'] = $scale;

                if (preg_match('{#?.*\?\/\?}', $format['Code'])) {
                    $format['Type'] = 'Fraction';
                } else {
                    $format['Code'] = str_replace('#', '', $format['Code']);

                    $matches = [];
                    if (preg_match('{(0+)(\.?)(0*)}', preg_replace('{\[[^\]]+\]}', '', $format['Code']), $matches)) {
                        $Integer = $matches[1];
                        $decimalPoint = $matches[2];
                        $decimals = $matches[3];

                        $format['MinWidth'] = strlen($Integer) + strlen($decimalPoint) + strlen($decimals);
                        $format['Decimals'] = $decimals;
                        $format['Precision'] = strlen($format['Decimals']);
                        $format['Pattern'] = '%0'.$format['MinWidth'].'.'.$format['Precision'].'f';
                    }
                }

                $matches = [];
                if (preg_match('{\[\$(.*)\]}u', $format['Code'], $matches)) {
                    $currFormat = $matches[0];
                    $currCode = $matches[1];
                    $currCode = explode('-', $currCode);
                    if ($currCode) {
                        $currCode = $currCode[0];
                    }

                    if (!$currCode) {
                        $currCode = $this->currencyCode;
                    }

                    $format['Currency'] = $currCode;
                }
                $format['Code'] = trim($format['Code']);
            }

            $this->parsedFormatCache[$index] = $format;
        }

        // Applying format to value
        if ($format) {
            if ($format['Code'] == '@') {
                return (string) $value;
            }
            // Percentages
            elseif ($format['Type'] == 'Percentage') {
                if ($format['Code'] === '0%') {
                    $value = round(100 * $value, 0).'%';
                } else {
                    $value = sprintf('%.2f%%', round(100 * $value, 2));
                }
            }
            // Dates and times
            elseif ($format['Type'] == 'DateTime') {
                $days = (int)$value;
                // Correcting for Feb 29, 1900
                if ($days > 60) {
                    $days--;
                }

                // At this point time is a fraction of a day
                $time = ($value - (int)$value);
                $seconds = 0;
                if ($time) {
                    // Here time is converted to seconds
                    // Some loss of precision will occur
                    $seconds = (int)($time * 86400);
                }

                $value = clone $this->baseDate;
                $value->add(new DateInterval('P'.$days.'D'.($seconds ? 'T'.$seconds.'S' : '')));

                if (!$this->option('returnDateTimeObjects')) {
                    $value = $value->format($format['Code']);
                } else {
                    // A DateTime object is returned
                }
            } elseif ($format['Type'] == 'Euro') {
                $value = 'EUR '.sprintf('%1.2f', $value);
            } else {
                // Fractional numbers
                if ($format['Type'] == 'Fraction' && ($value != (int)$value)) {
                    $integer = floor(abs($value));
                    $decimal = fmod(abs($value), 1);
                    // Removing the integer part and decimal point
                    $decimal *= pow(10, strlen($decimal) - 2);
                    $decimalDivisor = pow(10, strlen($decimal));

                    if ($this->GMPSupported) {
                        $GCD = gmp_strval(gmp_gcd($decimal, $decimalDivisor));
                    } else {
                        $GCD = self::GCD($decimal, $decimalDivisor);
                    }

                    $adjDecimal = $decimal/$GCD;
                    $adjDecimalDivisor = $decimalDivisor/$GCD;

                    if (
                        strpos($format['Code'], '0') !== false ||
                        strpos($format['Code'], '#') !== false ||
                        substr($format['Code'], 0, 3) == '? ?'
                    ) {
                        // The integer part is shown separately apart from the fraction
                        $value = ($value < 0 ? '-' : '').
                            $Integer ? $Integer.' ' : ''.
                            $adjDecimal.'/'.
                            $adjDecimalDivisor;
                    } else {
                        // The fraction includes the integer part
                        $adjDecimal += $integer * $adjDecimalDivisor;
                        $value = ($value < 0 ? '-' : '').
                            $adjDecimal.'/'.
                            $adjDecimalDivisor;
                    }
                } else {
                    // Scaling
                    $value = $value / $format['Scale'];

                    if (!empty($format['MinWidth']) && $format['Decimals']) {
                        if ($format['Thousands']) {
                            $value = number_format(
                                $value,
                                $format['Precision'],
                                $this->decimalSeparator,
                                $this->thousandSeparator
                            );
                        } else {
                            $value = sprintf($format['Pattern'], $value);
                        }

                        $value = preg_replace('{(0+)(\.?)(0*)}', $value, $format['Code']);
                    }
                }

                // Currency/Accounting
                if ($format['Currency']) {
                    $value = preg_replace('', $format['Currency'], $value);
                }
            }
        }

        return $value;
    }

    /**
     * Attempts to approximate Excel's "general" format.
     *
     * @param mixed value
     *
     * @return mixed
     */
    public function generalFormat($value)
    {
        // Numeric format
        if (is_numeric($value)) {
            // $value = (float)$value;
        }
        return $value;
    }

    /**
     * Takes the column letter and converts it to a numerical index (0-based)
     *
     * @param string letter(s) to convert
     *
     * @return mixed Numeric index (0-based) or boolean false if it cannot be calculated
     */
    public function indexFromColumnLetter($letter)
    {
        // $powers = [];
        $letter = strtoupper($letter);

        $result = 0;
        for ($i = strlen($letter) - 1, $j = 0; $i >= 0; $i--, $j++) {
            $ord = ord($letter[$i]) - 64;
            if ($ord > 26) {
                // Something is very, very wrong
                return false;
            }
            $result += $ord * pow(26, $j);
        }
        return $result - 1;
    }

    /**
     * Helper function for greatest common divisor calculation in case GMP extension is
     *	not enabled
     *
     * @param int Number #1
     * @param int Number #2
     *
     * @param int Greatest common divisor
     */
    public static function GCD($a, $b)
    {
        $a = abs($a);
        $b = abs($b);
        if ($a + $b == 0) {
            return 0;
        } else {
            $c = 1;
            while ($a > 0) {
                $c = $a;
                $a = $b % $a;
                $b = $c;
            }

            return $c;
        }
    }

    /**
     * Move forward to next element
     *
     * @return mixed
     */
    public function next()
    {
        $this->currentRowIndex++;

        if (!$this->rowOpen) {
            while ($this->valid = $this->worksheet->read()) {
                if ($this->worksheet->name == 'row') {
                    // Getting the row spanning area (stored as e.g., 1:12)
                    // so that the last cells will be present, even if empty
                    $rowSpans = $this->worksheet->getAttribute('spans');
                    if ($rowSpans) {
                        $rowSpans = explode(':', $rowSpans);
                        $currentRowColumnCount = $rowSpans[1];
                    } else {
                        $currentRowColumnCount = 0;
                    }

                    if ($currentRowColumnCount > 0) {
                        $this->currentRow = array_fill(0, $currentRowColumnCount, '');
                    }

                    $this->rowOpen = true;
                    break;
                }
            }
        }

        // Reading the necessary row, if found
        if ($this->rowOpen) {
            // These two are needed to control for empty cells
            $maxIndex = 0;
            $cellCount = 0;

            $cellHasSharedString = false;

            while ($this->valid = $this->worksheet->read()) {
                switch ($this->worksheet->name) {
                    // End of row
                    case 'row':
                        if ($this->worksheet->nodeType == XMLReader::END_ELEMENT) {
                            $this->rowOpen = false;
                            break 2;
                        }
                        break;
                    // Cell
                    case 'c':
                        // If it is a closing tag, skip it
                        if ($this->worksheet->nodeType == XMLReader::END_ELEMENT) {
                            break;
                        }

                        // Get the index of the cell
                        $index = $this->worksheet->getAttribute('r');
                        $letter = preg_replace('{[^[:alpha:]]}S', '', $index);
                        $index = $this->indexFromColumnLetter($letter);
                        // Get the style of the cell
                        $styleId = (int)$this->worksheet->getAttribute('s');

                        // Determine cell type
                        if ($this->worksheet->getAttribute('t') == self::CELL_TYPE_SHARED_STR) {
                            $cellHasSharedString = true;
                        } else {
                            $cellHasSharedString = false;
                        }

                        $this->currentRow[$index] = '';

                        $cellCount++;
                        if ($index > $maxIndex) {
                            $maxIndex = $index;
                        }

                        break;
                    // Cell value
                    case 'v':
                    case 'is':
                        if ($this->worksheet->nodeType != XMLReader::END_ELEMENT) {
                            $value = $this->worksheet->readString();
                            
                            if ($cellHasSharedString) {
                                $value = $this->getSharedString($value);
                            }
                           
                            // Format value if necessary
                            if ($value !== '' && $styleId && isset($this->stylesCache[$styleId])) {
                                $value = $this->formatValue($value, $styleId);
                            } elseif ($value) {
                                $value = $this->generalFormat($value);
                            }
    
                            $this->currentRow[$index] = $value;
                        }

                        break;
                }
            }

            // Adding empty cells, if necessary
            // Only empty cells inbetween and on the left side are added
            if ($maxIndex + 1 > $cellCount) {
                $this->currentRow = $this->currentRow + array_fill(0, $maxIndex + 1, '');
                ksort($this->currentRow);
            }
        }

        return $this->currentRow;
    }

    /**
     * Rewind the Iterator to the first element
     *
     * @return void
     */
    public function rewind()
    {
        // Removed the check whether $this->Index == 0 otherwise ChangeSheet doesn't work properly

        // If the worksheet was already iterated, XML file is reopened.
        // Otherwise it should be at the beginning anyway
        if ($this->worksheet instanceof XMLReader) {
            $this->worksheet->close();
        } else {
            $this->worksheet = new XMLReader;
        }

        $worksheetPath = $this->tempDir . 'xl/worksheets/sheet' . $this->sheetIndex . '.xml';

        $this->worksheet->open($worksheetPath);

        $this->valid = true;
        $this->rowOpen = false;
        
        parent::rewind();
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
     * Destructor
     */
    public function __destruct()
    {
        foreach ($this->tempFiles as $tempFile) {
            @unlink($tempFile);
        }

        // Better safe than sorry - shouldn't try deleting '.' or '/', or '..'.
        if (strlen($this->tempDir) > 2) {
            @rmdir($this->tempDir . 'xl' . DIRECTORY_SEPARATOR . 'worksheets');
            @rmdir($this->tempDir . 'xl');
            @rmdir($this->tempDir);
        }

        if ($this->worksheet && $this->worksheet instanceof XMLReader) {
            $this->worksheet->close();
            unset($this->worksheet);
        }

        if ($this->sharedStrings && $this->sharedStrings instanceof XMLReader) {
            $this->sharedStrings->close();
            unset($this->sharedStrings);
        }
        unset($this->sharedStringsPath);

        if (isset($this->stylesCache)) {
            unset($this->stylesCache);
        }
    }
}
