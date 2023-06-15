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

/**
 * Parse an object or array to json
 */
class Parser implements IParser
{
    /**
     * Store data to parse
     * @var DObject
     */
    private DObject $data;

    public function __construct(null|string $data = null)
    {
        if (!empty($data)) {
            $this->set($data);
        } else {
            $this->set('[]');
        }
    }

    /**
     * Set data to parser
     * @param string $data
     * @return Parser
     */
    public function set(string $data): Parser
    {
        $this->data = new DObject(json_decode($data));
        return $this;
    }

    /**
     * @param bool $json
     * @return false|string|DObject
     */
    public function get(bool $json = false): false|string|DObject
    {
        if ($json) {
            return $this->toJson();
        }

        return $this->data;
    }

    /**
     * Convert data to json
     * @return false|string
     */
    public function toJson(): false|string
    {
        return $this->data->toJSON();
    }

    /**
     * @return Parser
     */
    public function __invoke(): Parser
    {
        return $this;
    }

    /**
     * @return string
     */
    public function __toString(): string
    {
        return $this->toJson();
    }

    /**
     * @return Spread
     */
    public function toSpread(): Spread
    {
        return new Spread($this->data);
    }

    /**
     * @param string $string
     * @return bool
     */
    public function isJson(string $string): bool
    {
        json_decode($string);
        return (json_last_error() == JSON_ERROR_NONE);
    }
}
