<?php

class ExcelCreator
{

    public $excel;

    /**
     * @var PHPExcel_Worksheet
     */
    public $sheet;

    public $object;

    public $col = 0;

    public $cols;

    public $row = 1;

    public function __construct()
    {
        $this->excel = new \PHPExcel();
        $this->object = array();
        $this->cols =range('A', 'Z');
    }

    public function pushSingle($label, $value)
    {
        $this->sheet->setCellValue($this->getCoordinate(), $label);
        $this->row++;
        $this->sheet->setCellValue($this->getCoordinate(), $value);
        $this->col++;
        $this->row = 1;
    }

    private function getCoordinate()
    {
        return sprintf("%s%s",$this->getCol(), $this->row);
    }

    private function getCol()
    {
        $col = $this->cols[$this->col];
        return $col;
    }

    public function pushMultiple($label, array $values)
    {
        $this->sheet->setCellValue($this->getCoordinate(), $label);
        foreach ($values as $value) {
            $this->row++;
            $this->sheet->setCellValue($this->getCoordinate(), $value);
        }
        $this->row = 1;
    }

    public function createSheet($id, $title = null)
    {
        $this->row = 1;
        $this->col = 0;
        $this->excel->createSheet($id);
        $this->sheet = $this->excel->setActiveSheetIndex($id);
        $this->sheet->setTitle($title);
        $this->sheet->getStyle("A1")->getFont()->setBold();
    }

    public function setSheet()
    {
    }

    private function redimension()
    {
        $objPHPExcel = $this->excel;
        foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
            $objPHPExcel->setActiveSheetIndex($objPHPExcel->getIndex($worksheet));

            $sheet = $objPHPExcel->getActiveSheet();
            $cellIterator = $sheet->getRowIterator()->current()->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false);
            /** @var PHPExcel_Cell $cell */
            foreach ($cellIterator as $cell) {
                $sheet->getColumnDimension($cell->getColumn())->setAutoSize(true);
            }
        }
        $this->excel = $objPHPExcel;
    }

    public function getExcel()
    {
        $this->redimension();
        // We'll be outputting an excel file
//        header("Content-Encoding: UTF-8");
//        header('Content-type: application/vnd.ms-excel');

        // It will be called file.xls
//        header('Content-Disposition: attachment; filename="file.xls"');

        // Write file to the browser
        $file = 'file.xls';
        $objWriter = PHPExcel_IOFactory::createWriter($this->excel, 'Excel2007');
        $objWriter->save(__DIR__."/".$file);
    }
}