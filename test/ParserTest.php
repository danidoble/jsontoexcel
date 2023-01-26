<?php
/*
 * Copyright (c) 2023.
 * @author danidoble (Daniel Sandoval) <ddanidoble@gmail.com>
 * @website https://danidoble.com
 * @website https://github.com/danidoble
 */

namespace test;

use Danidoble\Jsontoexcel\Parser\Parser;
use PHPUnit\Framework\TestCase;

class ParserTest extends TestCase
{

    public function testSet()
    {
        $this->assertEquals(
            '["assigned"]',
            (new Parser(["assigned"]))
        );
    }

    public function testToJson()
    {
        $this->assertEquals(
            '{"dani":"doble"}',
            (new Parser(["dani" => "doble"]))->toJson(),
        );
    }

    public function testEmptyData()
    {
        $this->assertEquals(
            '[]',
            (new Parser())->toJson(),
        );
    }

    public function test__invoke()
    {
        $this->assertInstanceOf(
            Parser::class,
            new Parser()
        );
    }

    public function testGetArray()
    {
        $this->assertIsArray(
            (new Parser())->get()
        );
    }

    public function testGetObject()
    {
        $this->assertIsObject(
            (new Parser((object)[]))->get()
        );
    }
}
