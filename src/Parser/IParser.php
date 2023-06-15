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
     * @param string $data
     * @return Parser
     */
    public function set(string $data): Parser;

    /**
     * Get current data
     * @param bool $json
     * @return false|string|DObject
     */
    public function get(bool $json = false): false|string|DObject;

    /**
     * Convert data to json
     * @return false|string
     */
    public function toJson(): false|string;

    /**
     * @return Parser
     */
    public function __invoke(): Parser;

    /**
     * @return string
     */
    public function __toString(): string;

    /**
     * @return Spread
     */
    public function toSpread(): Spread;

    /**
     * @param string $string
     * @return bool
     */
    public function isJson(string $string): bool;
}
