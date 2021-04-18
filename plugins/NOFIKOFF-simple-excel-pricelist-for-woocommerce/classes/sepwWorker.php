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


    // если в админке не выбрано
    const DEFAULT_COLS = array('thumbnail', 'SKU', 'name', 'price', 'number', 'summ');

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
    private $current_row;
    private $kostyl_name_category_before;
    private $kostyl_id_product_unique;
    private $kostyl_category_not_empty;
    private $kostyl_category_not_empty_last_product_row = 7;



    private $max_number_columns;
    // TODO автоматически рассчитывать день
    private $last_number_columns = 'D';


    private $sheet;

    public function __construct()
    {
        parent::__construct();
        $this->max_number_columns = count($this->options['product_fields']);

        $this->worker = new ExcelWorker(dirname(__FILE__, 2));

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

        $this->fields = isset($this->options['product_fields']) ? $this->options['product_fields'] : self::DEFAULT_COLS;
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
        $this->sheet = $this->worker->sheet();
        //by Novikov 2019
        $this->sheet->setAutoFilter('A4:' . $this->last_number_columns . '4');


//        $products = wc_get_products(array(
//            'status' => 'publish',
//            'paginate' => false,
//            //'numberposts' => -1,
//            'numberposts' => 10,
//            'stock_status' => 'instock',
//        ));

        $this->head();
//        $this->body($this->sheet, $products);
        //by Novikov 2019
        $this->body();


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

    private function head()
    {
        $this->current_row = 1;

        //by Novikov 2019
        $cell = $this->sheet->getCellByColumnAndRow(1, $this->current_row);
        $coord = $cell->getCoordinate();
        $this->sheet->getStyle($coord)->getFont()->setBold(true);
        //$this->sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell->setValue('ТОВ "Меганом Україна" - кабель от производителя. Десятки тысяч наименований. Гибкий подход. Оперативная доставка. Дата формирования прайса: ' . date('d/m/Y'));

        for ($i = 1; $i <= $this->max_number_columns; $i++) {
            $cell = $this->sheet->getCellByColumnAndRow($i, $this->current_row);
            $coord = $cell->getCoordinate();
            $this->sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        }
        $this->current_row++;//2

        //by Novikov 2019
        $cell = $this->sheet->getCellByColumnAndRow(1, $this->current_row);
        $coord = $cell->getCoordinate();
        $this->sheet->getStyle($coord)->getFont()->setBold(true);
//        $this->sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell->setValue('б-р.Вацлава Гавела 8, м. Київ, +38 (044) 25 121 45, +38 (067) 56 54 402');

        for ($i = 1; $i <= $this->max_number_columns; $i++) {
            $cell = $this->sheet->getCellByColumnAndRow($i, $this->current_row);
            $coord = $cell->getCoordinate();
            $this->sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        }
        $this->current_row++;//3

        //by Novikov 2019
        $cell = $this->sheet->getCellByColumnAndRow(1, $this->current_row);
        $coord = $cell->getCoordinate();
        //$this->sheet->getStyle($coord)->getFont()->setBold(true);
        //$this->sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $cell->setValue('contact@meganom.kiev.ua');

        for ($i = 1; $i <= $this->max_number_columns; $i++) {
            $cell = $this->sheet->getCellByColumnAndRow($i, $this->current_row);
            $coord = $cell->getCoordinate();
            $this->sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        }
        $this->current_row++;//4

        // далее наполняем строку (кстати в начале метода - на эту строку вешаем эксель автофиьлтры)
        $col = 1;

        // выводим заголовки столбцев - это те у которых в slug head-xxxxxx
        foreach ($this->fields as $c) {
            foreach ($this->handlers as $h) {
                if ($h->fits('head-' . $c)) {
                    $h->write($this->sheet, $col, $this->current_row, NULL);
                    break;
                }
            }
            $col++;
        }
    }

    //by Novikov
    private function body()
    {

        // by Novikov 2019
        // начинаем печатать с этой строки
        $this->current_row = 5;


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
            'hide_empty' => $empty,
            //'child_of' => 0, //всех потомков без исключения для этого ID
            'parent' => 0, // только прямых потомков второго уровня для этого ID
        );
        $all_categories = get_categories($args);

        /** ПЕРЕБИРАЕМ ВСЕ КАТЕГОРИИ - берем первого уровня $cat->category_parent == 0 */
        foreach ($all_categories as $cat) {

            /* второй уровень */
            if ($cat->category_parent == 0) {
                $category_id = $cat->term_id;
                if ($cat->name !== 'Uncategorized') {
                    // title Category
                    $this->print_title_category($cat->name);
                }

                /* третий уровень TODO переписать на рекурсию без привязки к уровням */
                $args2 = array(
                    'taxonomy' => $taxonomy,
                    'child_of' => $category_id, //всех потомков без исключения для этого ID
                    //'parent' => 0, // только прямых потомков второго уровня для этого ID
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
                        //родительская и ниже второго уровня
                        $this->print_title_category($cat->name);
                        $this->print_title_category($sub_category->name);

                        if ($sub_category->slug === '') continue;
                        $this->print_row_products_by_cat_slug($sub_category->slug);
                    }
                }
            }

            if ($cat->slug === '') continue;
            $this->print_row_products_by_cat_slug($cat->slug);

            //            break;
        }

        // строка ИТОГО

        $cell = $this->sheet->getCellByColumnAndRow($this->max_number_columns - 2, $this->current_row);
        $coord = $cell->getCoordinate();
        $this->sheet->getStyle($coord)->getFont()->setBold(true);
        //$this->sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $this->sheet->getStyle($coord)->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_RIGHT);
        $this->sheet->getStyle($coord)->getAlignment()->setVertical(Style\Alignment::VERTICAL_CENTER);
        $cell->setValue('ИТОГО');

        $cell = $this->sheet->getCellByColumnAndRow($this->max_number_columns, $this->current_row);
        $coord = $cell->getCoordinate();
        $this->sheet->getStyle($coord)->getFont()->setBold(true);
        //$this->sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $this->sheet->getStyle($coord)->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_LEFT);
        $this->sheet->getStyle($coord)->getAlignment()->setVertical(Style\Alignment::VERTICAL_CENTER);
        $cell->setValue('=SUM(' . $this->last_number_columns . '' . $this->max_number_columns . ':' . $this->last_number_columns . '' . ($this->current_row - 1) . ')');

        for ($i = 1; $i <= $this->max_number_columns; $i++) {
            $cell = $this->sheet->getCellByColumnAndRow($i, $this->current_row);
            $coord = $cell->getCoordinate();
            $this->sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        }


    }

    private function print_title_category($cat_name)
    {
        // костыль
        if ($this->kostyl_name_category_before === $cat_name) return;

        $cell = $this->sheet->getCellByColumnAndRow(1, $this->current_row);
        $coord = $cell->getCoordinate();
        $this->sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');

        $cell = $this->sheet->getCellByColumnAndRow(2, $this->current_row);
        $coord = $cell->getCoordinate();
        $this->sheet->getStyle($coord)->getFont()->setBold(true);
        $this->sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        $this->sheet->getStyle($coord)->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_LEFT);
        $this->sheet->getStyle($coord)->getAlignment()->setVertical(Style\Alignment::VERTICAL_CENTER);
        $cell->setValue($cat_name);

        for ($i = 3; $i <= $this->max_number_columns; $i++) {
            $cell = $this->sheet->getCellByColumnAndRow($i, $this->current_row);
            $coord = $cell->getCoordinate();
            $this->sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FF66CC00');
        }
        $this->current_row++;
        $this->kostyl_name_category_before = $cat_name;
        $this->kostyl_category_not_empty = false;
    }

    private function print_row_products_by_cat_slug($category_slug)
    {
        //
        /** ДОСТАЕМ ПРОДУКТЫ ЭТОЙ КАТЕГОРИИ */
        $products = wc_get_products(
            [
                'status' => 'publish',
                'paginate' => false,
                'numberposts' => -1,
                //'numberposts' => 10,
                'category' => [$category_slug],
                'stock_status' => 'instock',
            ]);

        /** ПЕРЕБИРАЕМ ПРОДУКТЫ */
        foreach ($products as $p) {
            //
            if ($this->kostyl_id_product_unique[$p->get_id()]) continue;

            if ($p->is_type('variable')) {
                $this->simple_row($this->current_row, $p);
                $this->current_row++;
                $variations = $p->get_available_variations();
                foreach ($variations as $v) {
                    if ($v['variation_is_active'] && $v['variation_is_visible'] && $v['is_in_stock']) {
                        $this->variable_row($this->current_row, $p, $v);
                        $this->current_row++;

                        $this->kostyl_category_not_empty = true;
                        $this->kostyl_category_not_empty_last_product_row = $this->current_row;
                    }
                }

            } else {
                $this->simple_row($this->current_row, $p);
                $this->current_row++;

                $this->kostyl_category_not_empty = true;
                $this->kostyl_category_not_empty_last_product_row = $this->current_row;
            }

            $this->kostyl_id_product_unique[$p->get_id()] = true;
        }

        // эта категория пустая возвращаем голову на последний хвост когда товары еще были
        if (!$this->kostyl_category_not_empty) $this->current_row = $this->kostyl_category_not_empty_last_product_row;

    }

    private function simple_row($row, $product)
    {
        $col = 1;
        foreach ($this->fields as $c) {
            foreach ($this->handlers as $h) {
                if ($h->fits('simple-' . $c)) {
                    $h->write($this->sheet, $col, $row, $product);
                    break;
                }
            }
            $col++;
        }
    }

    private function variable_row($row, $product, $variation)
    {
        $col = 1;
        foreach ($this->fields as $c) {
            foreach ($this->handlers as $h) {
                if ($h->fits('var-' . $c)) {
                    $h->write($this->sheet, $col, $row, ['product' => $product, 'variation' => $variation]);
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
