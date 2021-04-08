<?php

namespace tests\ExcelTests;

use Chen\Excel\Readers\Csv;
use Chen\Excel\Readers\Ods;
use Chen\Excel\Readers\Xls;
use Chen\Excel\Readers\Xlsx;
use Chen\Excel\SpreadsheetReader;
use PHPUnit\Framework\TestCase;

class SpreadsheetReaderTest extends TestCase
{
    public function testLoadReader()
    {
        $filename = "tests/data/sample.csv";
        $reader = SpreadsheetReader::load($filename);
        $this->assertTrue($reader instanceof Csv);

        $filename = "tests/data/sample.ods";
        $reader = SpreadsheetReader::load($filename);
        $this->assertTrue($reader instanceof Ods);

        $filename = "tests/data/sample.xls";
        $reader = SpreadsheetReader::load($filename);
        $this->assertTrue($reader instanceof Xls);

        $filename = "tests/data/sample.xlsx";
        $reader = SpreadsheetReader::load($filename);
        $this->assertTrue($reader instanceof Xlsx);
    }

    public function testRegisterReader()
    {
        SpreadsheetReader::registerReader('Csv', SampleReader::class);
        $filename = "tests/data/sample.csv";
        $reader = SpreadsheetReader::load($filename);  
        $this->assertTrue($reader instanceof SampleReader);
        $this->assertTrue($reader->isReadable($filename));
    }
}
