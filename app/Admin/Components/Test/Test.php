<?php
namespace App\Admin\Components\Test;


class Test extends Base
{
    public function tabTitle()
    {
        return __('word生成工具');
    }

    /**
     * Build a form here.
     */
    public function form()
    {


        $this->file('doc1', '模板1')->required();
        $this->file('doc2', '模板2')->required();
        $this->file('data', '数据')->required();
        $this->multipleFile('images', '图片');


    }

}
