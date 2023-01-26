<?php
/*
 * Copyright (c) 2023.
 * @author danidoble (Daniel Sandoval) <ddanidoble@gmail.com>
 * @website https://danidoble.com
 * @website https://github.com/danidoble
 */

namespace Danidoble\Jsontoexcel\Parser;

/**
 * Parse an object or array to json
 */
class Parser implements IParser
{
    /**
     * Store data to parse
     * @var array|object
     */
    private array|object $data;

    public function __construct(null|array|object $data = [])
    {
        $this->set($data);
    }

    /**
     * Set data to parser
     * @param array|object $data
     * @return Parser
     */
    public function set(array|object $data = []): Parser
    {
        $this->data = $data;
        return $this;
    }

    /**
     * @param bool $json
     * @return false|array|object
     */
    public function get(bool $json = false): false|array|object
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
        return json_encode($this->data);
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
}
