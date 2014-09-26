<?php

namespace Andou\Exceller;

use \PHPExcel;
use \PHPExcel_Style_Fill;
use \PHPExcel_Style_Border;
use \PHPExcel_Style_Alignment;
use \PHPExcel_Writer_CSV;
use \PHPExcel_Writer_Excel2007;

/**
 * Your own personal Api Fetcher.
 * 
 * The MIT License (MIT)
 * 
 * Copyright (c) 2014 Antonio Pastorino <antonio.pastorino@gmail.com>
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 * 
 * 
 * @author Antonio Pastorino <antonio.pastorino@gmail.com>
 * @category exceller
 * @package andou/exceller
 * @copyright MIT License (http://opensource.org/licenses/MIT)
 * @todo Add documentation!!!!
 */
class Exceller {

  protected $_letters;
  protected $_objExcel;
  protected $_type = self::TYPE_XLS;
  protected $_save_path;
  protected $_file_name;
  protected $_creator;
  protected $_title;
  protected $_subject;
  protected $_active_sheet_title;

  const DS = '/';
  const TYPE_CSV = 'csv';
  const TYPE_XLS = 'xls';

  public function __construct() {
    $this->_letters = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');
    $this->_objExcel = new PHPExcel();
  }

  public function finalize() {
    $this->_saveExcel();
  }

  public function setType($type) {
    $this->_type = $type;
    return $this;
  }

  public function setSavePath($save_path) {
    $this->_save_path = rtrim($save_path, self::DS) . self::DS;
    return $this;
  }

  public function setFileName($file_name) {
    $this->_file_name = ltrim($file_name, self::DS);
    return $this;
  }

  public function setCreator($creator) {
    $this->_creator = $creator;
    return $this;
  }

  public function setTitle($title) {
    $this->_title = $title;
    return $this;
  }

  public function setSubject($subject) {
    $this->_subject = $subject;
    return $this;
  }

  public function setActiveSheetTitle($active_sheet_title) {
    $this->_active_sheet_title = $active_sheet_title;
    return $this;
  }

  /**
   * Inserisce il titolo del file excel
   */
  protected function _initExcel() {

    if ($this->_creator) {
      $this->_objExcel->getProperties()->setCreator($this->_creator);
      $this->_objExcel->getProperties()->setLastModifiedBy($this->_creator);
    }
    if ($this->_title) {
      $this->_objExcel->getProperties()->setTitle($this->_title);
    }
    if ($this->_subject) {
      $this->_objExcel->getProperties()->setSubject($this->_subject);
    }
    if ($this->_description) {
      $this->_objExcel->getProperties()->setDescription($this->_description);
    }
    return $this;
  }

  protected function _saveExcel() {
    if ($this->_active_sheet_title) {
      $this->_objExcel->getActiveSheet()->setTitle($this->_active_sheet_title);
    }
    switch ($this->_type) {
      case self::TYPE_CSV:
        $objWriter = new PHPExcel_Writer_CSV($this->_objExcel);
        break;
      case self::TYPE_XLS:
      default:
        $objWriter = new PHPExcel_Writer_Excel2007($this->_objExcel);
        break;
    }

    $objWriter->save($this->_composeSavePath());
  }

  protected function _composeSavePath() {
    return $this->_save_path . $this->_file_name . "." . $this->_getExtension();
  }

  protected function _getExtension() {
    switch ($this->_type) {
      case self::TYPE_CSV:
        return 'csv';
      default:
        return 'xlsx';
    }
  }

  /**
   * Inserisce una cella nel file excel
   * 
   * @param string $letter
   * @param int $number
   * @param string $value
   * @param array $style
   */
  public function insertCell($letter, $number, $value, $style = null) {

    if (is_int($letter)) {
      $letter = $this->getLetter($letter);
    }
    if (!is_null($style)) {
      $this->_objExcel
              ->getActiveSheet()
              ->SetCellValue(sprintf("%s%d", $letter, $number), $value)
              ->getStyle(sprintf("%s%d", $letter, $number))->applyFromArray($this->_getStyle($style));
    } else {
      $this->_objExcel
              ->getActiveSheet()
              ->SetCellValue(sprintf("%s%d", $letter, $number), $value);
    }
    return $this;
  }

  /**
   * Inserisce una cella di header
   * 
   * @param string $letter
   * @param string $number
   * @param string $value
   */
  public function insertHeaderCell($letter, $number, $value, $style = NULL) {
    $this->insertCell($letter, $number, $value, is_null($style) ? "HEADER" : $style);
    return $this;
  }

  protected function _getStyle($type) {
    $style = array();
    switch ($type) {
      case 'HEADER':
        $style = array(
            'font' => array(
                'name' => 'Calibri',
                'size' => 11,
                'bold' => true,
                'color' => array(
                    'rgb' => '000000'
                )
            ),
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => '7CCD7C'),
            ),
            'borders' => array(
                'outline' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('argb' => '000000'),
                ),
            ),
            'alignment' => array(
                'wrapText' => true,
                'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            )
        );
        break;
    }
    return $style;
  }

  /**
   * Restituisce una lettera sulla base di un count di colonna
   * 
   * @param int $cnt
   * @return string
   */
  public function getLetter($cnt) {

    $repeat = (int) ($cnt / count($this->_letters));
    $module = (int) (($cnt) % count($this->_letters));

    if ($repeat > 0) {
      return $this->_letters[$repeat - 1] . $this->_letters[$module];
    } else {
      return $this->_letters[$module];
    }
  }

}
