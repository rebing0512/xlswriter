<?php

namespace Jenson\Xlswriter\Helper;

use Jenson\Xlswriter\Service\XlswriterService;

class Helpers
{
    /**
     * Excel导出
     *
     * @param $path
     * @param $title
     * @param array $data
     * @param array $setting
     * @return void
     * @throws \Exception
     */
    public static function ExcelExport($path, $title,array $data, array $setting = array())
    {
        $xlswriter = new XlswriterService($path);
        $xlswriter->export($title,$data,$setting);
    }

    /**
     * Excel导入
     *
     * @param string $tale
     * @param $filePath
     * @param $filename
     * @param array $insert_field
     * @param int $setSkipRows
     * @return mixed
     * @throws \Exception
     */
    public static function ExcelImport($table,$filePath, $filename, array $insert_field, int $setSkipRows = 0)
    {
        $xlswriter = new XlswriterService($filePath);
        return $xlswriter->import($table,$filePath,$filename,$insert_field,$setSkipRows);
    }

    /**
     * 上传文件
     *
     * @param $parmas
     * @return array
     */
    public static function FileUpload($parmas): array
    {
        $xlswriter = new XlswriterService();
        $file = $parmas['file'];
        $fileExtra = [
            'file_size'   => $file['size'],
            'file_suffix' => $file['type'],
            'file_name'   => $file['name'],
        ];
        $save_path=$parmas['save_path']??'static/upload/files';
        return $xlswriter->fileUpload($file, $fileExtra,$save_path);
    }

    /**
     * 文件下载
     *
     * @param $file_url
     * @return void
     * @throws \Exception
     */
    public static function FileDownload($file_url){
        #文件路径
        if(empty($file_url)){
            throw new \Exception('文件不存在');
        }
        header("Content-Type: application/octet-stream");
        header("Content-Disposition: attachment; filename=" . basename($file_url));
        header("Content-Length: " . filesize($file_url));
        readfile($file_url);
    }

}