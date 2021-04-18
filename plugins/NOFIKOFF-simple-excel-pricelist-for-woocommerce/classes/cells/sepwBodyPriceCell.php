<?php

use PhpOffice\PhpSpreadsheet\Style;

class sepwBodyPriceCell extends sepwBodyCell
{
    /**
     * @param $sheet
     * @param int $col
     * @param int $row
     * @param WC_Product $p
     */
    public function write($sheet, $col, $row, $p)
    {

        switch ($this->slug) {
            case 'simple-price':
            case 'var-price':
                $price = html_entity_decode(strip_tags($p->get_price_html()));
                $price = trim(str_replace(" грн", "", $price));
                if (preg_match('/([\d.]+).([\d]{2})/', $price, $m)) {
                    $price = str_replace('.', '', $m[1]) . ',' . $m[2];
                } else {
                    $price = 0;
                }
                break;
            case 'simple-summ':
            case 'var-summ':
                $price = '=D' . $row . '*E' . $row;
                break;
            case 'simple-number':
            case 'var-number':
                $price = '';
                break;
        }


        $cell = $sheet->getCellByColumnAndRow($col, $row);
        $coord = $cell->getCoordinate();


        $cell->setValue($price);


        $sheet->getStyle($coord)->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($coord)->getAlignment()->setVertical(Style\Alignment::VERTICAL_CENTER);
        if ($p->is_type('variable')) {
            $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFEEEEEE');
        }
        $sheet->getStyle($coord)->getBorders()->getBottom()->setBorderStyle(Style\Border::BORDER_THIN);
        $sheet->getStyle($coord)->getBorders()->getRight()->setBorderStyle(Style\Border::BORDER_THIN);
    }
}