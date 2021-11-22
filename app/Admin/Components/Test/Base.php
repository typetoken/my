<?php

namespace App\Admin\Components\Test;

use Encore\Admin\Widgets\Form;
use Illuminate\Http\Request;
use Illuminate\Http\UploadedFile;
use Illuminate\Support\Facades\Cache;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpWord\TemplateProcessor;

abstract class Base extends Form
{
    /**
     * The form title.
     *
     * @var string
     */
    public $title;

    public function __construct($data = [])
    {
        parent::__construct($data);

        $this->title = $this->tabTitle();
    }


    abstract public function tabTitle();

    /**
     * author: mtg
     * time: 2021/1/21   11:25
     * function description: 子类可以自定义模块名
     * @return |null
     */
    public function module()
    {
        return null;

    }

    /**
     * Handle the form request.
     *
     * @param Request $request
     *
     * @return \Illuminate\Http\RedirectResponse
     */
    public function handle(Request $request)
    {
        ini_set("memory_limit","512M");
        ini_set("max_execution_time", "0");
        $keyValues = $request->all();
//        dd($keyValues);
        if (!$this->preHandle()) {
            return back();
        }
        $a = $this->dataChange($keyValues);

        dd($a);

        return back();

    }


    public function dataChange($keyValues)
    {
        $a = [];
        $imageNames = [];
        foreach ($keyValues as $key => $v) {
            if ($v instanceof UploadedFile) {
                $dataNewName = $key . '.' . request()->file($key)->getClientOriginalExtension();
                $path = Storage::disk('admin')->putFileAs('files', $v, $dataNewName);
                array_push($a, $path);
            }
            if ($key == 'images') {
                foreach ($v as $item) {
                    // 获取图片名字
                    $imageName = $item->getClientOriginalName();
                    $name = explode('.', $imageName)[0];
                    $imageNames[$name] = $imageName;
                    Storage::disk('admin')->putFileAs('images', $item, $imageName);
                }
            }
        }


        // execl
        $execl_path = storage_path('app/public') . '/' . $a[2];
        $execl_path = rtrim($execl_path, '\\');
        $inputFileType = \PHPExcel_IOFactory::identify($execl_path);
        $objReader = \PHPExcel_IOFactory::createReader($inputFileType);
        $execl = $objReader->load($execl_path);
        $sheet = $execl->getSheet(0);
        $highestRow = $sheet->getHighestRow();
//        $highestColumn = $sheet->getHighestColumn($highestRow);

        $elsxOneName = [];
        $alphabet = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
        // 获取表格首行文字数组
        foreach ($alphabet as $str) {
            $name = $sheet->rangeToArray($str . '1')[0][0];
            if (!$name) {
                continue;
            }
            $elsxOneName[$name] = $str;
        }

        // word1
        $doc1 = storage_path('app/public') . '/' . $a[0];
        $doc2 = storage_path('app/public') . '/' . $a[1];
        dump($highestRow);
        // 获取excel文件的数据，$row=2代表从第二行开始获取数据
        for ($row = 2; $row <= $highestRow; $row++){
            $word = new TemplateProcessor($doc1);
            $word2 = new TemplateProcessor($doc2);
            dump($row);
            foreach ($elsxOneName as $k => $name) {
                $value = $sheet->getCell($name . $row)->getValue();
                $word->setValue($k, $value);
                $word2->setValue($k, $value);
            }
            dump(333333);
            $imagePath = storage_path('app/public') . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR;
            // 找当前此行的用户名字
            $username = $sheet->getCell($elsxOneName['姓名'] . $row)->getValue();
            // 检查该用户是否有照片
            if (isset($imageNames[$username]) && file_exists($imagePath . $imageNames[$username])) {
                $word->setImageValue('照片', ['path'=>$imagePath . $imageNames[$username], 'width' => 110, 'height' => 140, 'ratio' => false]);
                $word2->setImageValue('照片', ['path'=>$imagePath . $imageNames[$username], 'width' => 110, 'height' => 140, 'ratio' => false]);
            }
            dump(44444);
            $pathName = (int)$sheet->getCell('A' . $row)->getValue();
            if (!is_dir(storage_path('app/public') . '/doc/'. $pathName) && mkdir(storage_path('app/public') . '/doc/'. $pathName, 0777, true)) {
                Log::channel('common')->info('权限问题');
                return back('权限问题');
            }
            dump(5555555555555);
            dump(storage_path('app/public') . DIRECTORY_SEPARATOR . 'doc' . DIRECTORY_SEPARATOR . $pathName . DIRECTORY_SEPARATOR .'doc1.docx');
            $word->saveAs(storage_path('app/public') . DIRECTORY_SEPARATOR . 'doc' . DIRECTORY_SEPARATOR . $pathName . DIRECTORY_SEPARATOR .'doc1.docx');
            $word2->saveAs(storage_path('app/public') . DIRECTORY_SEPARATOR . 'doc' . DIRECTORY_SEPARATOR . $pathName . DIRECTORY_SEPARATOR .'doc2.docx');
            dump(666666);

        }

        return 11111;

    }



    public function preHandle()
    {
        return true;
    }

    /**
     * author: mtg
     * time: 2021/1/21   11:21
     * function description: 配置的模块,即配置的类名
     * @return string
     */
    public function getModule()
    {
        $className = str_replace('\\', '/', get_called_class());

        return $this->module() ?? strtolower(basename($className));
    }

    /**
     * The data of the form.
     *
     *
     */
    public function data()
    {

        $data = [];


        return $data;
    }
}
