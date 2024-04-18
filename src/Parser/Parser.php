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
     */
    private DObject $data;

    public function __construct(?string $data = null)
    {
        if (! empty($data)) {
            $this->set($data);
        } else {
            $this->set('[]');
        }
    }

    /**
     * Set data to parser
     */
    public function set(string $data): Parser
    {
        $this->data = new DObject(json_decode($data));

        return $this;
    }

    public function get(bool $json = false): false|string|DObject
    {
        if ($json) {
            return $this->toJson();
        }

        return $this->data;
    }

    /**
     * Convert data to json
     */
    public function toJson(): false|string
    {
        return $this->data->toJSON();
    }

    public function __invoke(): Parser
    {
        return $this;
    }

    public function __toString(): string
    {
        return $this->toJson();
    }

    public function toSpread(): Spread
    {
        return new Spread($this->data);
    }

    public function isJson(string $string): bool
    {
        if (json_decode($string) !== null) {
            return true;
        }

        return false;
    }
}
