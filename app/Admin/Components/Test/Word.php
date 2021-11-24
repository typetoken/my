<?php
namespace App\Admin\Components\Test;


class Word extends Base
{
    public function tabTitle()
    {
        return __('word合成工具');
    }

    /**
     * Build a form here.
     */
    public function form()
    {


        $this->hidden('word_compound')->default(1);
        $this->text('lon_lat', '插入图片放置坐标')->help('例 100, 100  坐标将左右偏差部分');
        $this->file('image_template', '插入图片');
        $this->multipleFile('words', '模板图片');


    }

}
