<?php
/*
 * Copyright (c) 2023.
 * @author danidoble (Daniel Sandoval) <ddanidoble@gmail.com>
 * @website https://danidoble.com
 * @website https://github.com/danidoble
 */

namespace Danidoble\Jsontoexcel\Parser;

use Danidoble\DObject;
use Danidoble\Jsontoexcel\Spreadsheet\Spread;

interface IParser
{
    /**
     * Set data to parser
     */
    public function set(string $data): Parser;

    /**
     * Get current data
     */
    public function get(bool $json = false): false|string|DObject;

    /**
     * Convert data to json
     */
    public function toJson(): false|string;

    public function __invoke(): Parser;

    public function __toString(): string;

    public function toSpread(): Spread;

    public function isJson(string $string): bool;
}
