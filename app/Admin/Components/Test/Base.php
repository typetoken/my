<?php

namespace App\Admin\Components\Test;

use Encore\Admin\Widgets\Form;
use Grafika\Grafika;
use Illuminate\Http\Request;
use Illuminate\Http\UploadedFile;
use Illuminate\Support\Facades\Cache;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpWord\IOFactory;
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

        // word文档合成
        if (isset($keyValues['word_compound'])) {
            $resPath = $this->imageCompoundWord($keyValues);
        }else {
            $resPath = $this->dataChange($keyValues);
        }
        $zip = new \ZipArchive();

        $zipUrl = storage_path('app/public');

        $file_name = $zipUrl . $resPath;
        if (file_exists($file_name)) {
            unlink($file_name);
        }
        if (!is_dir($zipUrl . '/down')) {
            mkdir($zipUrl . '/down', 0777, true);
        }
        if ($zip->open($file_name, \ZipArchive::CREATE) === TRUE) {
            $this->addFileToZip(explode('.', $file_name)[0], $zip); //调用方法，对要打包的根目录进行操作，并将ZipArchive的对象传递给方法
            $zip->close(); //关闭处理的zip文件
        }

        return redirect('/admin/down?path=' . $resPath);
    }


    /**
     * 将文件生成压缩文件
     * @param $path
     * @param $zip
     * @param string $sub_dir
     */
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


    /**
     * 将execl表生成多个word文档
     * @param $keyValues
     * @return string
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     * @throws \PhpOffice\PhpWord\Exception\CopyFileException
     * @throws \PhpOffice\PhpWord\Exception\CreateTemporaryFileException
     */
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

        return '/down/doc.zip';

    }


    public function imageCompoundWord($keyValues)
    {
        // 处理数据
        $coord = explode(',', $keyValues['lon_lat']);
        $width = $coord[0] ?? 0;
        $height = $coord[1] ?? 0;

        $images = [];
        // 插入图片处理
        $ext = request()->file('image_template')->getClientOriginalExtension();
        $dataNewName = request()->file('image_template')->getClientOriginalName();
        $imageName = explode('.', $dataNewName)[0];
        $path = Storage::disk('admin')->putFileAs('images/' . $imageName, $keyValues['image_template'], $dataNewName);

        $spath = storage_path('app/public');

        // 图片旋转处理
        switch ($ext) {
            case 'jpeg':
                $source = imagecreatefromjpeg($spath . '/' . $path);

                // 生成新图
                for ($i = 0; $i < 30; $i++) {
                    // 获取旋转度数
                    $degrees = $i * 2;
                    $degrees1 = 360 - $degrees;
                    // 旋转
                    $rotate = imagerotate($source, $degrees, imageColorAllocateAlpha($source, 255, 255, 255, 127));
                    $rotate1 = imagerotate($source, $degrees1, imageColorAllocateAlpha($source, 255, 255, 255, 127));

                    $newPath = $spath . '/images/' . $imageName . '/' . $imageName . $degrees . '.' . $ext;
                    $newPath1 = $spath . '/images/' . $imageName . '/' . $imageName . $degrees1 . '.' . $ext;
                    // 生成新图
                    imagejpeg($rotate, $newPath);
                    imagejpeg($rotate1, $newPath1);
                    array_push($images, $newPath);
                    array_push($images, $newPath1);
                }

                break;
            case 'png':
                $source = imagecreatefrompng($spath . '/' . $path);

                // 生成新图
                for ($i = 0; $i < 30; $i++) {
                    // 获取旋转度数
                    $degrees = $i * 2;
                    $degrees1 = 360 - $degrees;
                    // 旋转
                    $rotate = imagerotate($source, $degrees, imageColorAllocateAlpha($source, 255, 255, 255, 127));
                    $rotate1 = imagerotate($source, $degrees1, imageColorAllocateAlpha($source, 255, 255, 255, 127));

                    $newPath = $spath . '/images/' . $imageName . '/' . $imageName . $degrees . '.' . $ext;
                    $newPath1 = $spath . '/images/' . $imageName . '/' . $imageName . $degrees1 . '.' . $ext;
                    // 生成新图
                    imagepng($rotate, $newPath);
                    imagepng($rotate1, $newPath1);
                    array_push($images, $newPath);
                    array_push($images, $newPath1);
                }

                break;
            default:
                Log::channel('common')->info('图片错误');
                return back();
        }

        // 处理图片模板
        foreach ($keyValues['words'] as $word) {
            $name = $word->getClientOriginalName();
            $imagePath = Storage::disk('admin')->putFileAs('words', $word, $name);
            // 为每个模板图片在坐标位置左右  添加上插入图片  合成一张新图片


            $width = mt_rand($width - 10, $width + 10);
            $height = mt_rand($height - 3, $height + 5);

            $this->imageEdit($spath . '/' . $imagePath, $images[array_rand($images)], $width, $height, $name);


        }

        return '/words/zip.zip';

    }


    public function imageEdit($backdrop, $insertImage, $width, $height, $name)
    {
        // 实例化图像编辑器
        $editor = Grafika::createEditor(['Gd']);
        // 打开背景图
        $editor->open($backImage, $backdrop);
        // 打开插入图片
        $editor->open($addImage, $insertImage);
        // 调整插入图片尺寸   暂不调整
//        $editor->resizeExact($avatarImage, '', '');
        // 图片插入到背景图片上
        $editor->blend($backImage, $addImage, 'normal', 1.0, 'top-left', $width, $height);

        // 保存图片
        $editor->save($backImage, storage_path('app/public') . '/words/zip/' . $name);
    }



    public function preHandle()
    {
        return true;
    }


    public function data()
    {

        $data = [];

        return $data;
    }
}
