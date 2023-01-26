<?php
/*
 * Copyright (c) 2023.
 * @author danidoble (Daniel Sandoval) <ddanidoble@gmail.com>
 * @website https://danidoble.com
 * @website https://github.com/danidoble
 */

namespace Danidoble\Jsontoexcel\Parser;

interface IParser
{
    /**
     * Set data to parser
     * @param array|object $data
     * @return Parser
     */
    public function set(array|object $data): Parser;

    /**
     * Get current data
     * @param bool $json
     * @return false|array|object
     */
    public function get(bool $json = false): false|array|object;

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
}
