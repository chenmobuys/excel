<?php

namespace tests\ExcelTests;

use Chen\Excel\Readers\BaseReader;

class SampleReader extends BaseReader
{
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
        return true;
    }
}
