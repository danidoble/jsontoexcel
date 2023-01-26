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

class Excel
{
    private Spreadsheet $spreadsheet;
    private Worksheet $sheet;
    private string $file_name = 'spreadsheet';
    private string $file_extension = 'xlsx';

    private array $keys = [];

    public function __construct(DObject $info)
    {
        $this->spreadsheet = new Spreadsheet();
        $this->sheet = $this->spreadsheet->getActiveSheet();
        $this->setDataToFile($info);
        //$spreadsheet = new Spreadsheet();
        //$sheet = $spreadsheet->getActiveSheet();
        //$sheet->setCellValue('A1', 'Hello World !');
        //
        //$writer = new Xlsx($spreadsheet);
        //$writer->save('hello world.xlsx');
    }

    public function setDataToFile(DObject $info): void
    {
        $this->setKeys($info[0]);
        $count_keys = count($this->keys);
        for ($i = 0; $i < $count_keys; $i++) {
            $this->sheet->setCellValue($this->getColumn($i) . '1', $this->keys[$i]);
        }

        $row = 2;
        foreach ($info as $item => $data) {
            for ($i = 0; $i < $count_keys; $i++) {
                $this->sheet->setCellValue($this->getColumn($i) . $row, $data[$this->keys[$i]]);
            }
            $row++;
        }
    }

    private function getColumn($n): string
    {
        for ($r = ""; $n >= 0; $n = intval($n / 26) - 1)
            $r = chr($n % 26 + 0x41) . $r;
        return $r;
    }

    public function setKeys($info): Excel
    {
        $this->keys = [];
        foreach ($info as $key => $ignore) {
            $this->keys[] = $key;
        }
        return $this;
    }

    public function setFileName(string $name): Excel
    {
        $this->file_name = $name;
        return $this;
    }

    public function getFileName(): string
    {
        return $this->file_name;
    }

    public function getExtension(): string
    {
        return $this->file_extension;
    }

    public function getFullFileName(): string
    {
        return urlencode($this->getFileName() . '.' . $this->getExtension());
    }

    public function setTypeFile(string $extension = 'xlsx'): Excel
    {
        $this->setExtension($extension);
        return $this;
    }

    /**
     * @throws Exception
     */
    public function getFile(): void
    {
        $writer = $this->getWriter();
        switch ($this->file_extension) {
            case 'csv':
                header("Content-Type: text/csv");
                break;
//            case 'ods':
//                header("Content-Type: application/vnd.oasis.opendocument.spreadsheet");
//                break;
//            case 'xls':
//            case 'xlsx':
//                header("Content-Type: application/vnd.ms-excel");
//                break;
            default:
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                break;
        }

        header('Content-Disposition: attachment; filename="' . $this->getFullFileName() . '"');
        $writer->save('php://output');
    }

    /**
     * @throws Exception
     */
    public function save(): void
    {
        $writer = $this->getWriter();
        $writer->save($this->getFullFileName());
    }

    private function setExtension(string $extension): void
    {
        $this->file_extension = strtolower($extension);
    }

    private function getWriter(): Ods|Csv|Xlsx|Xls
    {
        return match ($this->file_extension) {
            'xls' => $this->writerXls(),
            'ods' => $this->writerOds(),
            'csv' => $this->writerCsv(),
            default => $this->writerXlsx(),
        };
    }

    private function writerXlsx(): Xlsx
    {
        return new Xlsx($this->spreadsheet);
    }

    private function writerXls(): Xls
    {
        return new Xls($this->spreadsheet);
    }

    private function writerOds(): Ods
    {
        return new Ods($this->spreadsheet);
    }

    private function writerCsv(): Csv
    {
        return new Csv($this->spreadsheet);
    }
}