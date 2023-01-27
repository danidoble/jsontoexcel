<?php
/*
 * Copyright (c) 2023.
 * @author danidoble (Daniel Sandoval) <ddanidoble@gmail.com>
 * @website https://danidoble.com
 * @website https://github.com/danidoble
 */

namespace Danidoble\Jsontoexcel\Spreadsheet;

use Danidoble\DObject;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Writer\Ods;

class Spread implements ISpread
{
    /**
     * @var Spreadsheet current book
     */
    private Spreadsheet $spreadsheet;
    /**
     * @var Worksheet current sheet
     */
    private Worksheet $sheet;
    /**
     * @var string file name for download or save
     */
    private string $file_name = 'spreadsheet';
    /**
     * @var string extension to save or download file
     */
    private string $file_extension = 'xlsx';
    /**
     * @var array names of columns
     */
    private array $keys = [];

    /**
     * Constructor
     * @param DObject $info data to set into Excel
     */
    public function __construct(DObject $info)
    {
        $this->spreadsheet = new Spreadsheet();
        $this->sheet = $this->spreadsheet->getActiveSheet();
        $this->setDataToFile($info);
    }

    /**
     * Add new sheet to book
     * @param string $title title of new sheet
     * @param int|null $index index of new sheet
     * @return $this
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function addSheet(string $title, ?int $index = null): Spread
    {
        $sheet = new Worksheet($this->spreadsheet, $title);
        $this->spreadsheet->addSheet($sheet, $index);
        return $this;
    }

    /**
     * How many sheets are inside book
     * @return int
     */
    public function getSheetCount(): int
    {
        return $this->spreadsheet->getSheetCount();
    }

    /**
     * Get sheets names inside book
     * @return array
     */
    public function getSheetNames(): array
    {
        return $this->spreadsheet->getSheetNames();
    }

    /**
     * Set active worksheet by number (index)
     * @param int $index number of sheet (0,1...n)
     * @return $this
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setActiveSheet(int $index = 0): Spread
    {
        $this->sheet = $this->spreadsheet->getSheet($index);
        return $this;
    }

    /**
     * Set active worksheet by name of sheet
     * @param string $name name of the sheet
     * @return $this
     */
    public function setActiveSheetByName(string $name): Spread
    {
        $this->sheet = $this->spreadsheet->getSheetByName($name);
        return $this;
    }

    /**
     * Set name to current sheet
     * @param string $name
     * @return $this
     */
    public function setSheetName(string $name): Spread
    {
        $this->sheet->setTitle($name);
        return $this;
    }

    /**
     * Set data to file
     * @param DObject $info data to put inside sheet book
     * @return void
     */
    public function setDataToFile(DObject $info): void
    {
        $this->setKeys($info[0]);
        $count_keys = count($this->keys);
        for ($i = 0; $i < $count_keys; $i++) {
            $this->sheet->setCellValue($this->getColumn($i) . '1', $this->keys[$i]);
        }

        $row = 2;
        foreach ($info as $data) {
            for ($i = 0; $i < $count_keys; $i++) {
                $this->sheet->setCellValue($this->getColumn($i) . $row, $data[$this->keys[$i]]);
            }
            $row++;
        }
    }

    /**
     * Set column names
     * @param DObject $info 1st data to get names
     * @return $this
     */
    public function setKeys(DObject $info): Spread
    {
        $this->keys = [];
        foreach ($info as $key => $ignore) {
            $this->keys[] = $key;
        }
        return $this;
    }

    /**
     * Set file name
     * @param string $name
     * @return $this
     */
    public function setFileName(string $name): Spread
    {
        $this->file_name = $name;
        return $this;
    }

    /**
     * get current file name
     * @return string
     */
    public function getFileName(): string
    {
        return $this->file_name;
    }

    /**
     * get current extension
     * @return string
     */
    public function getExtension(): string
    {
        return $this->file_extension;
    }

    /**
     * get full current file name
     * @return string
     */
    public function getFullFileName(): string
    {
        return urlencode($this->getFileName() . '.' . $this->getExtension());
    }

    /**
     * set type file (extension)
     * @param string $extension
     * @return $this
     */
    public function setTypeFile(string $extension = 'xlsx'): Spread
    {
        $this->setExtension($extension);
        return $this;
    }

    /**
     * get file (download without save it)
     * @return void
     * @throws Exception
     */
    public function getFile(): void
    {
        $writer = $this->getWriter();
        switch ($this->file_extension) {
            case 'csv':
                header("Content-Type: text/csv");
                break;
            default:
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                break;
        }

        header('Content-Disposition: attachment; filename="' . $this->getFullFileName() . '"');
        $writer->save('php://output');
    }

    /**
     * save file
     * @param string|null $path path to save file
     * @return void
     * @throws Exception
     */
    public function save(?string $path = null): void
    {
        $writer = $this->getWriter();
        $writer->save(rtrim($path, '/\\').DIRECTORY_SEPARATOR.$this->getFullFileName());
    }

    /**
     * Get name of column
     * @param int $index get name column (0 = A, 3 = C)
     * @return string
     */
    private function getColumn(int $index): string
    {
        for ($r = ""; $index >= 0; $index = intval($index / 26) - 1) {
            $r = chr($index % 26 + 0x41) . $r;
        }
        return $r;
    }

    /**
     * Assign extension of file
     * @param string $extension extension type
     * @return void
     */
    private function setExtension(string $extension): void
    {
        $this->file_extension = strtolower($extension);
    }

    /**
     * make writer with right extension type
     * @return Ods|Csv|Xlsx|Xls
     */
    private function getWriter(): Ods|Csv|Xlsx|Xls
    {
        return match ($this->file_extension) {
            'xls' => $this->writerXls(),
            'ods' => $this->writerOds(),
            'csv' => $this->writerCsv(),
            default => $this->writerXlsx(),
        };
    }

    /**
     * Get Writer of xlsx
     * @return Xlsx
     */
    private function writerXlsx(): Xlsx
    {
        return new Xlsx($this->spreadsheet);
    }

    /**
     * Get Writer of xls
     * @return Xls
     */
    private function writerXls(): Xls
    {
        return new Xls($this->spreadsheet);
    }

    /**
     * Get Writer of ods
     * @return Ods
     */
    private function writerOds(): Ods
    {
        return new Ods($this->spreadsheet);
    }

    /**
     * Get Writer of csv
     * @return Csv
     */
    private function writerCsv(): Csv
    {
        return new Csv($this->spreadsheet);
    }
}
