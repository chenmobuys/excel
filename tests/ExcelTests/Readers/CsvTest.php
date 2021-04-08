<?php

namespace tests\ExcelTests\Readers;

use Chen\Excel\SpreadsheetReader;
use PHPUnit\Framework\TestCase;

class CsvTest extends TestCase
{
    public function testRead()
    {
        $filename = 'tests/data/sample.csv';
        $reader = SpreadsheetReader::load($filename);
        $sheetNamesExpected = ['sample.csv'];
        $rowsExpected = [
            ['Title','Description', 'Author'],
            ['This is title', 'This is description', 'This is author'],
        ];
        $rowsActual = [];
        foreach ($reader as $row) {
            $rowsActual[] = $row;
        }
        
        $this->assertEquals($sheetNamesExpected, $reader->getSheetNames());
        $this->assertEquals($rowsExpected, $rowsActual);
        $this->assertTrue(is_numeric($reader->count()));
    }
}
