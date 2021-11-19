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
        $this->title = __('推荐关系图');
    }


    public function index(Content $content)
    {
        $forms = [
            'test'    => Test::class,
        ];



        return $content
            ->title(__('配置信息'))
            ->body(Tab::forms($forms));
    }


}
