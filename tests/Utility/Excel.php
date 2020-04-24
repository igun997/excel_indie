<?php
namespace Indie\Utility;
require_once "Bootstrap.php";
use PhpOffice\PhpSpreadsheet\IOFactory;



/**
 * Excel Class
 */
class Excel
{
  CONST UP = 0;
  CONST DOWN = 1;
  public int $row = -1;
  public int $col = -1;
  public array $options;
  public $content;
  public $label;

  public ?string $type = null;

  function __construct($data,Array $options = [])
  {

    $spreadsheet = IOFactory::load($data);

    $this->options = $options;
    $this->content = $spreadsheet;
    return $this;
  }

  public function setLabel(int $row,int $col = -1)
  {
    $this->row = $row;
    $this->col = $col;
    return $this;
  }


  public function fromUrl(String $url)
  {
    return $this;
  }

  public function fromBase64(String $data)
  {
    return $this;
  }

  public function fromFilesRequest($data)
  {
    return $this;
  }

  public function type(String $type)
  {

    $this->type = $type;
    if ($this->type == "raw") {


    }elseif ($this->type == "json") {


    }elseif ($this->type == "xml") {


    }elseif ($this->type == "array") {

      $this->content =  $this->arrayFormat();
    }else {
      $this->content =  false;
    }
    return $this;
  }

  private function arrayFormat()
  {

    $sheetData = $this->content->getActiveSheet()->toArray(null, true, true, true);
    $this->_replacer($sheetData);
    return $sheetData;

  }

  private function _replacer(&$data)
  {
    $i = 0;
    if (is_array($data)) {
      foreach ($data as $key => &$value) {
        if (!is_numeric($key)) {
          unset($data[$key]);
          $data[$i] = $value;
        }else {
          $this->_replacer($value);
        }
        $i++;
      }
    }

  }

  private function _index($data,$val)
  {
    foreach ($data as $key => $value) {
      if ($val == $value) {
        return $key;
      }
    }
    throw new \Exception('index data notfound');
  }

  public function reformat(Array $options = [])
  {

    $use_col = true;

    if (empty($options)) {
      $options = $this->options;
    }

    if (empty($options)) {

      throw new \Exception('Reformat Need Options Params on `reformat()` or on initial construct');
      exit();
    }



    if ($this->row === -1) {

      throw new \Exception('setLabel must be contruct first');
      exit();
    }



    if ($this->col === -1) {
      $use_col = false;
    }

    if (!$use_col) {
      if (!isset($this->content[$this->row])) {

        throw new \Exception('label not found on index '.$this->row);
        exit();

      }

      $this->label = $this->content[$this->row];
      $startFrom = [$this->row+1];

    }else {

      if (!isset($this->content[$this->col][$this->row])) {
        throw new \Exception('label not found on index '.$this->row.' - '.$this->col);
        exit();

      }

      $this->label = $this->content[$this->col][$this->row];
      $startFrom = [($this->col),$this->row+1];

    }

    $in_array = [];

    foreach ($this->label as $key => $value) {
      $in_array[] = strtolower($value);
    }





    if (count($startFrom) == 1) {

       $startFormat = $this->content[$startFrom[0]];
    }else {

       $startFormat = $this->content[$startFrom[0]][$startFrom[1]];
    }

    $build = [];

    foreach ($options as $k => $v) {
      if (!is_array($v)) {

        if (in_array(strtolower($v),$in_array)) {
          $build[$k] = $startFormat[$this->_index($in_array,strtolower($v))];
        }else {
          $build[$k] = NULL;
        }
      }else {

        foreach ($v as $key => &$value) {
          if (in_array(strtolower($value),$in_array)) {
            $build[$k][$key] = $startFormat[$this->_index($in_array,strtolower($value))];
          }else {
            $build[$k][$key] = NULL;
          }
        }
      }


    }

    $this->label = $build;

    return $this;

  }

  public function output()
  {

    return $this->label;

  }

}
