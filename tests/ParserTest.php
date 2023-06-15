<?php
/*
 * Copyright (c) 2023.
 * @author danidoble (Daniel Sandoval) <ddanidoble@gmail.com>
 * @website https://danidoble.com
 * @website https://github.com/danidoble
 */

namespace Tests;

use Danidoble\DObject;
use Danidoble\Jsontoexcel\Parser\Parser;
use PhpOffice\PhpSpreadsheet\Worksheet\Protection;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PHPUnit\Framework\TestCase;

final class ParserTest extends TestCase
{
    public function testIsJson(): void
    {
        $parser = new Parser();
        $this->assertTrue($parser->isJson('{"name":"danidoble"}'));
    }

    public function testFailIsJson(): void
    {
        $parser = new Parser();
        $this->assertTrue($parser->isJson('"name":"danidoble"}'));
    }

    /**
     * @throws Exception
     */
    public function testToExcelFile(): void
    {
        $json = '[{"id":1,"first_name":"Janka","last_name":"Muzzlewhite","email":"jmuzzlewhite0@state.tx.us","gender":"Female","ip_address":"159.251.192.185"},{"id":2,"first_name":"Garrek","last_name":"Yarnton","email":"gyarnton1@google.com","gender":"Male","ip_address":"139.116.97.9"},{"id":3,"first_name":"Tyrone","last_name":"Losseljong","email":"tlosseljong2@symantec.com","gender":"Male","ip_address":"112.16.146.178"}]';
        $spread = new Parser($json);
        $spread = $spread->toSpread();
        $spread->setSheetName('phpunit_test');
        $spread->setFileName('phpUnitTest');
        $spread->setKeys(new DObject(['# ID', 'Nombre(s)', 'Apellido(s)', 'Correo', 'Genero', 'Direccion IP']));
        $spread->save(__DIR__);
        $this->assertFileExists(__DIR__.'/phpUnitTest.xlsx');
        if(file_exists(__DIR__.'/phpUnitTest.xlsx')) {
            unlink(__DIR__.'/phpUnitTest.xlsx');
        }
        $this->assertFileDoesNotExist(__DIR__.'/phpUnitTest.xlsx');
    }
}
