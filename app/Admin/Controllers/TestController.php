<?php

namespace App\Admin\Controllers;

use App\Admin\Components\Test\Test;
use Encore\Admin\Controllers\AdminController;
use Encore\Admin\Layout\Content;
use Encore\Admin\Widgets\Tab;

class TestController extends AdminController
{

    protected $title;


    public function __construct()
    {
        $this->title = __('工具');
    }


    public function index(Content $content)
    {
        $forms = [
            'test'    => Test::class,
        ];



        return $content
            ->title(__('工具'))
            ->body(Tab::forms($forms));
    }


    public function down(Content $content)
    {
        $downPath = $_SERVER['HTTP_HOST'] . '/storage/down/doc.zip';

        return $content
            ->title('Dashboard')
            ->description('Description...')
            ->row('下载地址: <a href="/storage/down/doc.zip" target="view_window" >' . $downPath . '</a>')
            ->row('请复制到另一个网页打开下载');
    }


}
