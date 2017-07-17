<?php
/**
 * Created by PhpStorm.
 * User: lenovo
 * Date: 2016/11/3
 * Time: 18:14
 */

namespace common\helpers;

use yii\base\Object;

class ExcelHelper extends Object
{

    const SUMMARY_TEXT="TEXT";
    const SUMMARY_SUM="SUM";

    private $index = 1;

    private $_bodyBeginIndex;
    private $_bodyEndIndex;


    private $title = [];
    private $info = [];

    private $headers = [];

    private $sheets = [];

    public function renderFile($fileName)
    {
        $file = $this->createFile();
        ob_end_clean();
        ob_start();
        header('Content-Type : application/vnd.ms-excel');
        header('Content-Disposition:attachment;filename="'.$fileName.'"');
        $file->save('php://output');
    }

    public function setTitle($title)
    {
        $this->title = $title;
    }
    public function setInfo($info)
    {
        $this->info = $info;
    }
    public function setHeaders($headers)
    {
        $this->headers = $headers;
    }

    public function setSheets($sheets)
    {
        $this->sheets = $sheets;
    }

    public function saveFile($path)
    {
        $file = $this->createFile();
        $file->save($path);
    }

    private function createFile()
    {
        $objectPHPExcel = new \PHPExcel();
        foreach ($this->sheets as $index=>$sheet)
        {
            $activeSheet  = $objectPHPExcel->setActiveSheetIndex($index);
            if(!empty($sheet["name"]))
            {
                $activeSheet->setTitle($sheet["name"]);
            }
            $this->renderHeader($activeSheet);

            $this->renderBody($activeSheet,$sheet["data"]);
            $this->renderSummary($activeSheet);
        }
        $objWriter= \PHPExcel_IOFactory::createWriter($objectPHPExcel,'Excel5');
        return $objWriter;
    }

    private function renderHeader(&$activeSheet)
    {
        $activeSheet->getPageSetup()->setHorizontalCentered(true);
        $activeSheet->getPageSetup()->setVerticalCentered(true);

        $count =  $this->getColsCount();
        $count = $count-1;
        $endCol = $this->getHeaderChar($count);
        $startCol = $this->getHeaderChar(0);

        //TITLE
        if($this->title)
        {
            $activeSheet->mergeCells($startCol.$this->index.':'.$endCol.$this->index);
            $activeSheet->setCellValue($startCol.$this->index,$this->title["text"]);
            $activeSheet->getStyle($startCol.$this->index)->getFont()->setSize(24);
            $activeSheet->getStyle($startCol.$this->index)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $this->index= $this->index+1;
        }
        //INFO
        foreach ($this->info as $key=>$value)
        {
            $col = $this->getHeaderChar($count-1);
            $activeSheet->setCellValue($col.$this->index,$key);
            $activeSheet->setCellValue($endCol.$this->index,$value);
            $activeSheet->getStyle($endCol.$this->index.':'.$endCol.$this->index)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
            $this->index = $this->index+1;
        }
        //HEADER

        $colIndex = 0;
        $maxHeaderDeep = 1;
        foreach ($this->headers as $i=>$header)
        {
            if(isset($header["headers"]))
            {
                list($colspan,$rowspan)=$this->renderGroupHeaderCell($activeSheet,$header,$colIndex,$this->index);
            }else
            {
                list($colspan,$rowspan)=$this->renderHeaderCell($activeSheet,$header,$colIndex,$this->index);
            }
            $colIndex = $colIndex+$colspan;

            if($rowspan>$maxHeaderDeep)
            {
                $maxHeaderDeep=$rowspan;
            }
        }

        $style = $activeSheet->getStyle($startCol.$this->index.':'.$endCol.($this->index+$maxHeaderDeep-1));
        $style->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $style->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
        //设置颜色
        $style->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CCCC');

        //设置边框
        $style->getBorders()->getAllBorders()->setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);
        $this->index = $this->index+$maxHeaderDeep;
    }

    private function renderHeaderCell(&$activeSheet,$header,$colIndex,$rowIndex)
    {

        $title =  $header["title"];
        $col = $this->getHeaderChar($colIndex);
        $activeSheet->setCellValue($col.$rowIndex,$title);

        //$colspan = isset($header["colspan"])?$header["colspan"]:1;
        $colspan=1;
        $rowspan = isset($header["rowspan"])?$header["rowspan"]:1;
        $endCol = $this->getHeaderChar($colIndex+$colspan-1);

        //合并单元格
        if($colspan>1 || $rowspan>1)
        {
            $activeSheet->mergeCells($col.$rowIndex.':'.$endCol.($rowIndex+$rowspan-1));
        }

        if(isset($header["width"]))
        {
            $activeSheet->getColumnDimension($col)->setWidth($header["width"]);
        }
        return [$colspan,$rowspan];
    }

    private function renderGroupHeaderCell(&$activeSheet,$header,$colIndex,$rowIndex)
    {
        $title =  $header["title"];
        $col = $this->getHeaderChar($colIndex);
        $activeSheet->setCellValue($col.$this->index,$title);
        $index = $rowIndex+1;
        $cIndex = $colIndex;
        $deep = 0;
        foreach ($header["headers"] as $i=>$h)
        {
            list($colspan,$rowspan) = $this->renderHeaderCell($activeSheet,$h,$cIndex,$index);
            $cIndex = $cIndex+$colspan;
            if($rowspan>$deep);
            {
                $deep=$rowspan;
            }
        }
        $rowspan = $deep;
        $colspan = ($cIndex-$colIndex)==0?1:($cIndex-$colIndex);

        $endCol = $this->getHeaderChar($colIndex+$colspan-1);
        //合并单元格
        $activeSheet->mergeCells($col.$rowIndex.':'.$endCol.($rowIndex+$rowspan-1));

        return [$colspan,$rowspan];
    }

    private  function renderBody(&$activeSheet,$data)
    {
        $this->_bodyBeginIndex = $this->index;
        foreach ($data as $d)
        {
            $colIndex = 0;
            foreach ($this->headers as $header)
            {
                if(isset($header["headers"]))
                {
                    foreach ($header["headers"] as $h)
                    {
                        $colspan=$this->renderContent($activeSheet,$h,$d,$colIndex);
                        $colIndex = $colIndex+$colspan;
                    }
                }else
                {
                    $colspan=$this->renderContent($activeSheet,$header,$d,$colIndex);
                    $colIndex = $colIndex+$colspan;
                }
            }
            $this->index = $this->index+1;
        }
        $this->_bodyEndIndex = $this->index-1;
    }

    private function renderContent(&$activeSheet,$header,$model,$colIndex)
    {
        $content = "";
        if(isset($header["render"]))
        {
            $content = $header["render"]($model);
        }else if(isset($header["data"]))
        {
            if(is_array($model))
            {
                $attrs = $model;
            }else
            {
                $attrs = $model->attributes;
            }
            if(isset($attrs[$header["data"]]))
            {
                $content =$attrs[$header["data"]];
            }
        }
        $col = $this->getHeaderChar($colIndex);
        $activeSheet->setCellValue($col.$this->index,$content);
        $activeSheet->getStyle($col.$this->index)->getAlignment()->setWrapText(true);
        return 1;
    }

    private  function renderSummary(&$activeSheet)
    {
        $colIndex = 0;
        foreach ($this->headers as $header)
        {
            if(isset($header["headers"]))
            {
                foreach ($header["headers"] as $h)
                {
                    if(isset($h["summary"]))
                    {
                        $col = $this->getHeaderChar($colIndex);
                        $content = $this->renderSummaryContent($h["summary"],$colIndex);
                        $activeSheet->setCellValue($col.$this->index, $content);
                    }
                    $colIndex = $colIndex+1;
                }
            }else
            {
                if(isset($header["summary"]))
                {
                    $col = $this->getHeaderChar($colIndex);
                    $content = $this->renderSummaryContent($header["summary"],$colIndex);
                    $activeSheet->setCellValue($col.$this->index, $content);
                }
                $colIndex = $colIndex+1;
            }
        }
        $this->index = $this->index+1;
    }

    private function renderSummaryContent($summary,$colIndex)
    {
        switch ($summary["type"])
        {
            case self::SUMMARY_TEXT:
            {
                return $summary["text"];
            }
            case self::SUMMARY_SUM:
            {
                $col = $this->getHeaderChar($colIndex);
                return '=SUM('.$col.$this->_bodyBeginIndex.':'.$col.$this->_bodyEndIndex.')';
            }
        }

    }

    private function getColsCount()
    {
        $colIndex = 0;
        foreach ($this->headers as $header)
        {
            if(isset($header["headers"]))
            {
                foreach ($header["headers"] as $h)
                {
                    $colIndex = $colIndex+1;
                }
            }else
            {
                $colIndex = $colIndex+1;
            }
        }
        return $colIndex;
    }

    private function getHeaderChar($index, $start = 65) {
        $str = '';
        if (floor($index / 26) > 0) {
            $str .= $this->getHeaderChar(floor($index / 26)-1);
        }
        return $str . chr($index % 26 + $start);
    }

}
