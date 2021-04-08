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
    public function load(string $filename, array $options = [])
    {
    }

    /**
     * Checks file is readable
     *
     * @param string $filename
     *
     * @return bool
     */
    public function isReadable(string $filename)
    {
        return true;
    }
}
