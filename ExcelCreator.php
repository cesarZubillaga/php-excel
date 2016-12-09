<?php

class ExcelCreator
{
    /**
     * @var \PHPExcel
     */
    public $excel;

    /**
     * @var \PHPExcel_Worksheet
     */
    public $sheet;

    /**
     * The index that points to the columns
     * @var int
     */
    public $col = 0;

    /**
     * Excel columns [A-ZZ]
     * @var array
     */
    public $cols;

    /**
     * The index that points to the active
     * @var int
     */
    public $row;

    /**
     * Variable used when is a horizontal display.
     * @var int
     */
    public $hRow;

    /**
     * ExcelCreator constructor.
     */
    public function __construct()
    {
        $this->excel = new \PHPExcel();
        $_ranges = range('A', 'Z');
        $ranges = $_ranges;
        for ($i = 0; $i < sizeof($_ranges); $i++) {
            foreach ($_ranges as $range) {
                $ranges[] = $_ranges[$i] . $range;
            }
        }
        $this->cols = $ranges;
        $this->printLabels = true;
        $this->hRow = 1;
        $this->row = 1;
    }

    /**
     * @param $label
     * @param $values
     * @return array
     */
    public static function getExportUnit($label, $values)
    {
        return array(
            'label' => $label,
            'values' => (is_array($values)) ? $values : array($values),
        );
    }

    /**
     * Push to a column multiple values
     * @param $label
     * @param array $values
     */
    public function pushV($label, array $values)
    {
        $this->sheet->setCellValue($this->getCoordinate(), $label);
        foreach ($values as $value) {
            $this->row++;
            $this->sheet->setCellValue($this->getCoordinate(), $value);
        }
        $this->col++;
        $this->row = 1;
    }

    /**
     * Push variables on the horizontal way
     * @param $label
     * @param array $values
     */
    public function pushH($label, array $values)
    {
        if ($this->printLabels) {
            $this->sheet->setCellValue($this->getCoordinate(), $label);
        }
        $this->row = $this->hRow;
        $this->row++;
        foreach ($values as $value) {
            $this->sheet->setCellValue($this->getCoordinate(), $value);
            $this->col++;
        }
        $this->row = $this->hRow;
    }

    /**
     * Insert a value when it's going to be k=>v format
     * @param $label
     * @param $value
     */
    public function singleH($label, $value)
    {
        if ($this->printLabels) {
            $this->sheet->setCellValue($this->getCoordinate(), $label);
        }
        $this->row = $this->hRow;
        $this->row++;
        $this->sheet->setCellValue($this->getCoordinate(), $value);
        $this->col++;
        $this->row = $this->hRow;
    }

    /**
     *
     */
    public function dontPrintLabels()
    {
        $this->printLabels = false;
    }

    /**
     * Get the column
     * @return array
     */
    private function getCol()
    {
        $col = $this->cols[$this->col];
        return $col;
    }

    /**
     * Resets the Column value
     */
    public function reset()
    {
        $this->col = 0;
        $this->hRow += 1;
    }

    /**
     * Create a Sheet with a title
     * @param $id
     * @param null $title
     */
    public function createSheet($id, $title = null)
    {
        $this->row = 1;
        $this->col = 0;
        $this->excel->createSheet($id);
        $title = sprintf("%s...", substr($title, 0, 28));
        $this->sheet = $this->excel->setActiveSheetIndex($id);
        $this->sheet->setTitle($title);
    }

    /**
     * Create the Excel Object
     * @param string $file
     */
    public function getBufferExcel($file = 'file_name')
    {
        $this->redimension();
        $this->excel->setActiveSheetIndex(0);
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header(sprintf('Content-Disposition: attachment;filename="%s.xls"', $file));
        header('Cache-Control: max-age=0');

        $objWriter = \PHPExcel_IOFactory::createWriter($this->excel, 'Excel5');
        $objWriter->save('php://output');
    }

    /**
     * Redimension all the columns of the Sheets
     */
    protected function redimension()
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

    /**
     * Get the active coordinate using the column and the row
     * @return string
     */
    protected function getCoordinate()
    {
        return sprintf("%s%s", $this->getCol(), $this->row);
    }

}