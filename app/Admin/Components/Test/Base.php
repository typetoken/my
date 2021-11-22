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
        $this->dataChange($keyValues);
        $zip = new \ZipArchive();

        $zipUrl = storage_path('app/public');

        $file_name = $zipUrl . '/down/doc.zip';
        if (file_exists($file_name)) {
            unlink($file_name);
        }
        if (!is_dir($zipUrl . '/down')) {
            mkdir($zipUrl . '/down', 0777, true);
        }
        if ($zip->open($file_name, \ZipArchive::CREATE) === TRUE) {
            $this->addFileToZip($zipUrl . '/doc/', $zip); //调用方法，对要打包的根目录进行操作，并将ZipArchive的对象传递给方法
            $zip->close(); //关闭处理的zip文件
        }

        return redirect('/admin/down');
    }


    public function addFileToZip($path, $zip, $sub_dir = '') {
        $handler = opendir($path); //打开当前文件夹由$path指定。
        /*
        循环的读取文件夹下的所有文件和文件夹
        其中$filename = readdir($handler)是每次循环的时候将读取的文件名赋值给$filename，
        为了不陷于死循环，所以还要让$filename !== false。
        一定要用!==，因为如果某个文件名如果叫'0'，或者某些被系统认为是代表false，用!=就会停止循环
        */
        while (($filename = readdir($handler)) !== false) {
            if ($filename != "." && $filename != "..") {//文件夹文件名字为'.'和‘..’，不要对他们进行操作
                if (is_dir($path . $filename)) {// 如果读取的某个对象是文件夹，则递归
                    $this->addFileToZip($path . "/" . $filename, $zip, $filename . '/');
                } else { //将文件加入zip对象
                    $zip->addFile($path . "/" . $filename, $sub_dir . $filename);
                }
            }
        }
        @closedir($path);
    }



    public function dataChange($keyValues)
    {
        // 源模板路径
        $a = [];
        // 源模板名字
        $fileNames = [];
        // 所有上传图片名字
        $imageNames = [];

        foreach ($keyValues as $key => $v) {
            if ($v instanceof UploadedFile) {
//                $dataNewName = $key . '.' . request()->file($key)->getClientOriginalExtension();
                $dataNewName = request()->file($key)->getClientOriginalName();
                $path = Storage::disk('admin')->putFileAs('files', $v, $dataNewName);
                array_push($fileNames, $dataNewName);
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

        // 删除掉doc
        deldir(storage_path('app/public') . '/doc/');

        // word1
        $doc1 = storage_path('app/public') . '/' . $a[0];
        $doc2 = storage_path('app/public') . '/' . $a[1];

        // 获取excel文件的数据，$row=2代表从第二行开始获取数据
        for ($row = 2; $row <= $highestRow; $row++){
            $word = new TemplateProcessor($doc1);
            $word2 = new TemplateProcessor($doc2);

            foreach ($elsxOneName as $k => $name) {
                $value = $sheet->getCell($name . $row)->getValue();
                $word->setValue($k, $value);
                $word2->setValue($k, $value);
            }

            $imagePath = storage_path('app/public') . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR;
            // 找当前此行的用户名字
            $username = $sheet->getCell($elsxOneName['姓名'] . $row)->getValue();
            // 检查该用户是否有照片
            if (isset($imageNames[$username]) && file_exists($imagePath . $imageNames[$username])) {
                $word->setImageValue('照片', ['path'=>$imagePath . $imageNames[$username], 'width' => 110, 'height' => 140, 'ratio' => false]);
                $word2->setImageValue('照片', ['path'=>$imagePath . $imageNames[$username], 'width' => 110, 'height' => 140, 'ratio' => false]);
            }

//            $pathName = (int)$sheet->getCell('A' . $row)->getValue();
//            $pathName = iconv('UTF-8','GBK',$username);

            $pathName = $username;
            Log::channel('common')->info(file_exists(storage_path('app/public') . '/doc/'. $pathName));
            if (!file_exists(storage_path('app/public') . '/doc/'. $pathName)) {
                mkdir(storage_path('app/public') . '/doc/'. $pathName, 0777, true);
            }
            chmod(storage_path('app/public') . '/doc/'. $pathName, 0777);

            $word->saveAs(storage_path('app/public') . DIRECTORY_SEPARATOR . 'doc' . DIRECTORY_SEPARATOR . $pathName . DIRECTORY_SEPARATOR .$fileNames[0]);
            $word2->saveAs(storage_path('app/public') . DIRECTORY_SEPARATOR . 'doc' . DIRECTORY_SEPARATOR . $pathName . DIRECTORY_SEPARATOR .$fileNames[1]);

        }

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
