<?php

namespace tests\ExcelTests\Readers;

use Chen\Excel\SpreadsheetReader;
use PHPUnit\Framework\TestCase;

class OdsTest extends TestCase
{
    public function testRead()
    {
        $filename = 'tests/data/sample.ods';
        $reader = SpreadsheetReader::load($filename);
        $sheetNamesExpected = ['工作表1', '工作表2', '工作表3'];
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
    
        $reader->setSheetIndex(1);
        $rowsExpected = [
            ['Title 2','Description 2', 'Author 2'],
            ['This is title', 'This is description', 'This is author'],
        ];
        $rowsActual = [];
        foreach ($reader as $row) {
            $rowsActual[] = $row;
        }

        $this->assertEquals($rowsExpected, $rowsActual);
        $this->assertTrue(is_numeric($reader->count()));
    }
}
