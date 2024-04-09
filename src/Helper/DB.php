<?php

namespace Jenson\Xlswriter\Helper;

use Medoo\Medoo;

class DB
{
    public $database;

    public function __construct()
    {
        $this->database = new medoo([
            // 必须配置项
            'database_type' => 'mysql',
            'database_name' => getenv('DB_DATABASE'),
            'server'        => getenv('DB_HOST'),
            'username'      => getenv('DB_USERNAME'),
            'password'      => getenv('DB_PASSWORD'),
            'charset'       => 'utf8',
            'port'          => getenv('DB_PORT'),
            'prefix'        => '',
            // 连接参数扩展, 更多参考 http://www.php.net/manual/en/pdo.setattribute.php
            'option' => [
                \PDO::ATTR_CASE => \PDO::CASE_NATURAL
            ]
        ]);
    }
}