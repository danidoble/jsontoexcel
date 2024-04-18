<?php

test('to excel file', function () {
    $json = '[{"id":1,"first_name":"Janka","last_name":"Muzzlewhite","email":"jmuzzlewhite0@state.tx.us","gender":"Female","ip_address":"159.251.192.185"},{"id":2,"first_name":"Garrek","last_name":"Yarnton","email":"gyarnton1@google.com","gender":"Male","ip_address":"139.116.97.9"},{"id":3,"first_name":"Tyrone","last_name":"Losseljong","email":"tlosseljong2@symantec.com","gender":"Male","ip_address":"112.16.146.178"}]';
    $spread = new \Danidoble\Jsontoexcel\Parser\Parser($json);
    $spread = $spread->toSpread();
    $spread->setSheetName('phpunit_test');
    $spread->setFileName('phpUnitTest');
    $spread->setKeys(new \Danidoble\DObject(['# ID', 'Nombre(s)', 'Apellido(s)', 'Correo', 'Genero', 'Direccion IP']));
    $spread->save(__DIR__);
    expect(__DIR__.'/phpUnitTest.xlsx')->toBeFile();
    if (file_exists(__DIR__.'/phpUnitTest.xlsx')) {
        unlink(__DIR__.'/phpUnitTest.xlsx');
    }
    $this->assertFileDoesNotExist(__DIR__.'/phpUnitTest.xlsx');
});
