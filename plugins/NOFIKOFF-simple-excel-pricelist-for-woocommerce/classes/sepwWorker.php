<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style;
use OnestExcelWriter\TextCell;

use OnestExcelWriter\Worker as ExcelWorker;

class sepwWorker extends sepwBootstrap
{

    const DEFAULT_COLS = array('thumbnail', 'SKU', 'name', 'price', 'stock');

    /**
     * @var array
     */
    private $handlers;

    /**
     * @var ExcelWorker
     */
    private $worker;

    /**
     * @var array
     */
    private $fields;

    public function __construct()
    {
        parent::__construct();


        $this->worker = new ExcelWorker(dirname(dirname(__FILE__)));

        $this->handlers = array(
            new sepwHeadNameCell('head-name'),
            new sepwHeadCell('head-SKU'),
            new sepwHeadCell('head-thumbnail'),
            new sepwHeadCell('head-price'),
            new sepwHeadCell('head-stock'),
            new sepwHeadCell('head-number'),
            new sepwHeadCell('head-summ'),

            new sepwBodyThumbnailCell('simple-thumbnail', $this->worker->getTmp()),
            new sepwBodyCell('simple-SKU'),
            new sepwBodyCell('simple-name'),
            new sepwBodyPriceCell('simple-price'),

            //by Novikov
            new sepwBodyPriceCell('simple-number'),
            new sepwBodyPriceCell('simple-summ'),

            //new sepwBodyXXXXXX что тчо было затер('simple-stock'),


            new sepwVarNameCell('var-name'),
            //new sepwVarSKUCell('var-SKU'),
            new sepwVarStockCell('var-stock'),
            //new sepwVarThumbnailCell('var-thumbnail', $this->worker->getTmp()),
            new sepwVarPriceCell('var-price'),

            //by Novikov
            new sepwVarPriceCell('var-number'),
            new sepwVarPriceCell('var-summ'),
        );

        add_action('rest_api_init', array($this, 'rest_api_init'));
        add_shortcode('pricelist', array($this, 'pricelist_shortcode'));

        //$this->fields = isset($this->options['product_fields']) ? $this->options['product_fields'] : self::DEFAULT_COLS;
        $this->fields = ['thumbnail', 'SKU', 'name', 'price', 'number', 'summ'];


    }

    public function generate_callback()
    {
        $this->generate();

        return array(
            'status' => true,
            'time' => date('d.m.Y H:i:s'),
        );
    }

    private function generate()
    {
        $sheet = $this->worker->sheet();
        //by Novikov 2019
        //авто фильтры
        $sheet->setAutoFilter('A4:F4');


//        $products = wc_get_products(array(
//            'status' => 'publish',
//            'paginate' => false,
//            //'numberposts' => -1,
//            'numberposts' => 10,
//            'stock_status' => 'instock',
//        ));

        $this->head($sheet);
//        $this->body($sheet, $products);
        //by Novikov 2019
        $this->body($sheet);


        //by Novikov 2019
        $this->worker->save('meganom.kiev.ua_pricelist.xlsx');

        update_option('sepw_generated', time());
    }

    public function rest_api_init()
    {
        register_rest_route('sepw/v1', '/generate', array(
            'methods' => 'GET',
            'callback' => array($this, 'generate_callback'),
        ));
    }

    private function head($sheet)
    {

        //by Novikov 2019
        $cell = $sheet->getCellByColumnAndRow(1, 1);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFont()->setBold(true);
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell->setValue('ТОВ "Меганом Україна" - кабель от производителя. Десятки тысяч наименований. Гибкий подход. Оперативная доставка<. Дата формирования прайса: ' . date('d/m/Y'));

        $cell = $sheet->getCellByColumnAndRow(2, 1);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell = $sheet->getCellByColumnAndRow(3, 1);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell = $sheet->getCellByColumnAndRow(4, 1);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell = $sheet->getCellByColumnAndRow(5, 1);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell = $sheet->getCellByColumnAndRow(6, 1);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');


        //by Novikov 2019
        $cell = $sheet->getCellByColumnAndRow(1, 2);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFont()->setBold(true);
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell->setValue('б-р.Вацлава Гавела 8, м. Київ, +38 (044) 25 121 45, +38 (067) 56 54 402');

        $cell = $sheet->getCellByColumnAndRow(2, 2);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell = $sheet->getCellByColumnAndRow(3, 2);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell = $sheet->getCellByColumnAndRow(4, 2);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell = $sheet->getCellByColumnAndRow(5, 2);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell = $sheet->getCellByColumnAndRow(6, 2);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');

        //by Novikov 2019
        $cell = $sheet->getCellByColumnAndRow(1, 3);
        $coord = $cell->getCoordinate();
        //$sheet->getStyle($coord)->getFont()->setBold(true);
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell->setValue('contact@meganom.kiev.ua');

        $cell = $sheet->getCellByColumnAndRow(2, 3);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell = $sheet->getCellByColumnAndRow(3, 3);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell = $sheet->getCellByColumnAndRow(4, 3);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell = $sheet->getCellByColumnAndRow(5, 3);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell = $sheet->getCellByColumnAndRow(6, 3);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');


        // далее наполняем строку (кстати в начале метода - на эту строку вешаем автофиьлтры)
        $col = 1;
        $row = 4;


        // выводим заголовки столбцев - это те у которых в slug head-xxxxxx
        foreach ($this->fields as $c) {
            foreach ($this->handlers as $h) {
                if ($h->fits('head-' . $c)) {
                    $h->write($sheet, $col, $row, NULL);
                    break;
                }
            }
            $col++;
        }
    }

    //by Novikov
    private function body($sheet)
    {

        // by Novikov 2019
        $row = 5;


        // ЗАДАЧА В ПРАЙСЕ СГРУППИРОВАТЬ ЦЕНЫ ПО КАТЕГОРИЯМ

        /** ДОСТАЕМ КАТЕГОРИИ */
        $taxonomy = 'product_cat';
        $orderby = 'name';
        $show_count = 0;      // 1 for yes, 0 for no
        $pad_counts = 0;      // 1 for yes, 0 for no
        $hierarchical = 1;      // 1 for yes, 0 for no
        $title = '';
        $empty = 0;

        $args = array(
            'taxonomy' => $taxonomy,
            'orderby' => $orderby,
            'show_count' => $show_count,
            'pad_counts' => $pad_counts,
            'hierarchical' => $hierarchical,
            'title_li' => $title,
            'hide_empty' => $empty
        );
        $all_categories = get_categories($args);

        /** ПЕРЕБИРАЕМ КАТЕГОРИИ */
        foreach ($all_categories as $cat) {
            $category_slug = '';
            if ($cat->category_parent == 0) {
                //$category_id = $cat->term_id;
                if ($cat->name != 'Uncategorized') {
                    $category_slug = $cat->slug;

                    $cell = $sheet->getCellByColumnAndRow(1, $row);
                    $coord = $cell->getCoordinate();
                    $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');

                    $cell = $sheet->getCellByColumnAndRow(2, $row);
                    $coord = $cell->getCoordinate();
                    $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');


                    $cell = $sheet->getCellByColumnAndRow(3, $row);
                    $coord = $cell->getCoordinate();
                    $sheet->getStyle($coord)->getFont()->setBold(true);
                    $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
                    $sheet->getStyle($coord)->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_LEFT);
                    $sheet->getStyle($coord)->getAlignment()->setVertical(Style\Alignment::VERTICAL_CENTER);
                    $cell->setValue($cat->name);

                    $cell = $sheet->getCellByColumnAndRow(4, $row);
                    $coord = $cell->getCoordinate();
                    $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');

                    $cell = $sheet->getCellByColumnAndRow(5, $row);
                    $coord = $cell->getCoordinate();
                    $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');

                    $cell = $sheet->getCellByColumnAndRow(6, $row);
                    $coord = $cell->getCoordinate();
                    $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');


                    $row++;
                }

                /* !!! ЕСЛИ ПОНАДОБЯТСЯ ПОДКАТЕГОРИИ
                $args2 = array(
                    'taxonomy' => $taxonomy,
                    'child_of' => 0,
                    'parent' => $category_id,
                    'orderby' => $orderby,
                    'show_count' => $show_count,
                    'pad_counts' => $pad_counts,
                    'hierarchical' => $hierarchical,
                    'title_li' => $title,
                    'hide_empty' => $empty
                );
                $sub_cats = get_categories($args2);
                if ($sub_cats) {
                    foreach ($sub_cats as $sub_category) {
                        echo $sub_category->name . " <br>\n";
                    }
                }*/
            }

            if ($category_slug == '') continue;
            //
            /** ДОСТАЕМ ПРОДУКТЫ ЭТОЙ КАТЕГОРИИ */
            $products = wc_get_products(
                [
                    'status' => 'publish',
                    'paginate' => false,
                    //'numberposts' => -1,
                    'numberposts' => 10,
                    'category' => [$category_slug],
                    'stock_status' => 'instock',
                ]);

            /** ПЕРЕБИРАЕМ ПРОДУКТЫ */
            foreach ($products as $p) {
                //
                if ($p->is_type('variable')) {
                    $this->simple_row($sheet, $row, $p);
                    $row++;
                    $variations = $p->get_available_variations();
                    foreach ($variations as $v) {
                        if ($v['variation_is_active'] && $v['variation_is_visible'] && $v['is_in_stock']) {
                            $this->variable_row($sheet, $row, $p, $v);
                            $row++;
                        }
                    }

                } else {
                    $this->simple_row($sheet, $row, $p);
                    $row++;
                }
            }
//            break;
        }

        // ИТОГО

        $cell = $sheet->getCellByColumnAndRow(7, $row);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFont()->setBold(true);
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $sheet->getStyle($coord)->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_LEFT);
        $sheet->getStyle($coord)->getAlignment()->setVertical(Style\Alignment::VERTICAL_CENTER);
        $cell->setValue('ИТОГО');

        $cell = $sheet->getCellByColumnAndRow(6, $row);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getFont()->setBold(true);
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $sheet->getStyle($coord)->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_LEFT);
        $sheet->getStyle($coord)->getAlignment()->setVertical(Style\Alignment::VERTICAL_CENTER);
        $cell->setValue('=SUM(F6:F' . ($row - 1) . ')');


    }

    private function simple_row($sheet, $row, $product)
    {
        $col = 1;
        foreach ($this->fields as $c) {
            foreach ($this->handlers as $h) {
                if ($h->fits('simple-' . $c)) {
                    $h->write($sheet, $col, $row, $product);
                    break;
                }
            }
            $col++;
        }
    }

    private function variable_row($sheet, $row, $product, $variation)
    {
        $col = 1;
        foreach ($this->fields as $c) {
            foreach ($this->handlers as $h) {
                if ($h->fits('var-' . $c)) {
                    $h->write($sheet, $col, $row, ['product' => $product, 'variation' => $variation]);
                    break;
                }
            }
            $col++;
        }
    }

    /**
     * @param array attrs
     * @return string
     */
    public function pricelist_shortcode($attrs)
    {
        $title = isset($attrs['title']) ? $attrs['title'] : __('Download Pricelist');
        $classes = isset($attrs['class']) ? $attrs['class'] : '';
        return '<a href="' . $this->pricelist_url() . '" class="' . $classes . '">' . $title . '</a>';
    }

    private function pricelist_url()
    {
        if (!$this->valid()) {
            $this->generate();
        }
        return $this->plugin_url . 'out/meganom.kiev.ua_pricelist.xlsx';
    }
}
