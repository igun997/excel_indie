<?php
require "loader.php";
use Indie\Utility\Debug;
use Indie\Utility\Excel;

$excel = new Excel(__DIR__."/example2.xls");

$options = [
  "name"=>"city",
  "country"=>"country",
  "lat"=>"latitude",
  "lng"=>"longitude",
];

$res = $excel->type("array")->setLabel(1)->reformat($options)->output();

Debug::log([$res]);
