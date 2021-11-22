<?php

function deldir($path){

    // 检查地址是否存在 是否是文件夹
    if (!is_dir($path)) {
        return true;
    }

    // 列出文件夹里的文件和目录
    $childPaths = scandir($path);

    foreach ($childPaths as $childPath) {
        // 如果是目录   递归
        if ($childPath != '.' && $childPath != '..') {
            if (is_dir($path . $childPath)) {
                deldir($path . $childPath . '/');
                // 而后删除文件夹
                rmdir($path . $childPath . '/');
            }else {
                // 否则   删除文件
                unlink($path . $childPath);
            }
        }
    }

    return true;
}
