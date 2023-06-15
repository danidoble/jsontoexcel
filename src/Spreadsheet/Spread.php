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
     * @var DObject names of columns
     */
    private DObject $keys;
    /**
     * @var DObject
     */
    private DObject $custom_keys;
    /**
     * @var DObject
     */
    private DObject $data;

    /**
     * Constructor
     * @param DObject $info data to set into Excel
     */
    public function __construct(DObject $info)
    {
        $this->spreadsheet = new Spreadsheet();
        $this->sheet = $this->spreadsheet->getActiveSheet();
        $this->data = $info;
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
        if (!isset($this->keys)) {
            $this->keys = new DObject([]);
            $k = $info->{'0'};
            if ($k === null) {
                $k = new DObject([]);
            }
            $this->setPrivateKeys($k);
        }
        $count_keys = $this->keys->count();
        if (!isset($this->custom_keys)) {
            $this->custom_keys = $this->keys;
        }
        for ($i = 0; $i < $count_keys; $i++) {
            $this->sheet->setCellValue($this->getColumn($i) . '1', $this->custom_keys[$i]);
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
     * @param DObject|array $info 1st data to get names
     * @return $this
     */
    public function setKeys(DObject|array $info): Spread
    {
        $keys = $this->getArrKeys($info);
        $this->custom_keys = new DObject($keys);
        return $this;
    }

    /**
     * Get current keys for titles
     * @return DObject
     */
    public function getKeys(): DObject
    {
        return $this->keys;
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
        $this->setDataToFile($this->data);
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
        $this->setDataToFile($this->data);
        $writer = $this->getWriter();
        $writer->save(rtrim($path, '/\\') . DIRECTORY_SEPARATOR . $this->getFullFileName());
    }

    /**
     * get working file
     * @return Worksheet
     */
    public function getWorkingFile(): Worksheet
    {
        return $this->sheet;
    }

    /**
     * set working file
     * @return $this
     */
    public function setWorkingFile(Worksheet $sheet): static
    {
        $this->sheet = $sheet;
        return $this;
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

    /**
     * Get keys from array
     * @param array|DObject $info
     * @return array
     */
    private function getArrKeys(array|DObject $info): array
    {
        $keys = [];
        if (!is_array($info)) {
            $info = $info->toArray();
        }
        $first_key = array_key_first($info);
        if (is_string($first_key)) {
            foreach ($info as $key => $ignore) {
                $keys[] = $key;
            }
        } else {
            foreach ($info as $key) {
                $keys[] = $key;
            }
        }
        return $keys;
    }

    /**
     * Set internal keys to work with, these keys are not able to mutate by user
     * @param DObject|array $info
     * @return void
     */
    private function setPrivateKeys(DObject|array $info): void
    {
        $keys = $this->getArrKeys($info);
        $this->keys = new DObject($keys);
    }
}
