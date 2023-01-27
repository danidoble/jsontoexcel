<?php
/*
 * Copyright (c) 2023.
 * @author danidoble (Daniel Sandoval) <ddanidoble@gmail.com>
 * @website https://danidoble.com
 * @website https://github.com/danidoble
 */

namespace Danidoble\Jsontoexcel\Spreadsheet;

use Danidoble\DObject;
use PhpOffice\PhpSpreadsheet\Writer\Exception;

interface ISpread
{
    /**
     * Constructor
     * @param DObject $info data to set into Excel
     */
    public function __construct(DObject $info);

    /**
     * Add new sheet to book
     * @param string $title title of new sheet
     * @param int|null $index index of new sheet
     * @return $this
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function addSheet(string $title, ?int $index = null): Spread;

    /**
     * How many sheets are inside book
     * @return int
     */
    public function getSheetCount(): int;

    /**
     * Get sheets names inside book
     * @return array
     */
    public function getSheetNames(): array;

    /**
     * Set active worksheet by number (index)
     * @param int $index number of sheet (0,1...n)
     * @return $this
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setActiveSheet(int $index = 0): Spread;

    /**
     * Set active worksheet by name of sheet
     * @param string $name name of the sheet
     * @return $this
     */
    public function setActiveSheetByName(string $name): Spread;

    /**
     * Set name to current sheet
     * @param string $name
     * @return $this
     */
    public function setSheetName(string $name): Spread;

    /**
     * Set data to file
     * @param DObject $info data to put inside sheet book
     * @return void
     */
    public function setDataToFile(DObject $info): void;

    /**
     * Set column names
     * @param DObject $info 1st data to get names
     * @return $this
     */
    public function setKeys(DObject $info): Spread;

    /**
     * Set file name
     * @param string $name
     * @return $this
     */
    public function setFileName(string $name): Spread;

    /**
     * get current file name
     * @return string
     */
    public function getFileName(): string;

    /**
     * get current extension
     * @return string
     */
    public function getExtension(): string;

    /**
     * get full current file name
     * @return string
     */
    public function getFullFileName(): string;

    /**
     * set type file (extension)
     * @param string $extension
     * @return $this
     */
    public function setTypeFile(string $extension = 'xlsx'): Spread;

    /**
     * get file (download without save it)
     * @return void
     * @throws Exception
     */
    public function getFile(): void;

    /**
     * save file
     * @param string|null $path path to save file
     * @return void
     * @throws Exception
     */
    public function save(?string $path = null): void;
}
