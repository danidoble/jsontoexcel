<?php
/*
 * Copyright (c) 2023.
 * @author danidoble (Daniel Sandoval) <ddanidoble@gmail.com>
 * @website https://danidoble.com
 * @website https://github.com/danidoble
 */

namespace Danidoble\Jsontoexcel\Spreadsheet;

use Danidoble\DObject;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;

interface ISpread
{
    /**
     * Constructor
     *
     * @param  DObject  $info  data to set into Excel
     */
    public function __construct(DObject $info);

    /**
     * Add new sheet to book
     *
     * @param  string  $title  title of new sheet
     * @param  int|null  $index  index of new sheet
     * @return $this
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function addSheet(string $title, ?int $index = null): Spread;

    /**
     * How many sheets are inside book
     */
    public function getSheetCount(): int;

    /**
     * Get sheets names inside book
     */
    public function getSheetNames(): array;

    /**
     * Set active worksheet by number (index)
     *
     * @param  int  $index  number of sheet (0,1...n)
     * @return $this
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setActiveSheet(int $index = 0): Spread;

    /**
     * Set active worksheet by name of sheet
     *
     * @param  string  $name  name of the sheet
     * @return $this
     */
    public function setActiveSheetByName(string $name): Spread;

    /**
     * Set name to current sheet
     *
     * @return $this
     */
    public function setSheetName(string $name): Spread;

    /**
     * Set data to file
     *
     * @param  DObject  $info  data to put inside sheet book
     */
    public function setDataToFile(DObject $info): void;

    /**
     * Set column names
     *
     * @param  DObject|array  $info  1st data to get names
     * @return $this
     */
    public function setKeys(DObject|array $info): Spread;

    /**
     * Get current keys for titles
     */
    public function getKeys(): DObject;

    /**
     * Set file name
     *
     * @return $this
     */
    public function setFileName(string $name): Spread;

    /**
     * get current file name
     */
    public function getFileName(): string;

    /**
     * get current extension
     */
    public function getExtension(): string;

    /**
     * get full current file name
     */
    public function getFullFileName(): string;

    /**
     * set type file (extension)
     *
     * @return $this
     */
    public function setTypeFile(string $extension = 'xlsx'): Spread;

    /**
     * get file (download without save it)
     *
     * @throws Exception
     */
    public function getFile(): void;

    /**
     * save file
     *
     * @param  string|null  $path  path to save file
     *
     * @throws Exception
     */
    public function save(?string $path = null): void;

    public function getWorkingFile(): Worksheet;

    /**
     * @return $this
     */
    public function setWorkingFile(Worksheet $sheet): static;

    public function setData(DObject $info): Spread;
}
