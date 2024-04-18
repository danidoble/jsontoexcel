<?php

test('is json', function () {
    $parser = new \Danidoble\Jsontoexcel\Parser\Parser();
    expect($parser->isJson('{"name":"danidoble"}'))->toBeTrue();
});

test('fail is json', function () {
    $parser = new \Danidoble\Jsontoexcel\Parser\Parser();
    expect($parser->isJson('"name":"danidoble"}'))->toBeFalse();
});
