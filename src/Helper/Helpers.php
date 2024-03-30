<?php

namespace Jenson\Xlswriter\Helper;

use Jenson\Xlswriter\Service\XlswriterService;

class Helpers
{
    /**
     * Excel导出
     *
     * @param $title
     * @param $data
     * @param $setting
     * @return void
     * @throws \Exception
     */
    public static function ExcelExport($title,$data,$setting = array())
    {
        $xlswriter = new XlswriterService();
        $xlswriter::export($title,$data,$setting);
    }

    /**
     * Excel导入
     *
     * @param $filePath
     * @param $filename
     * @param array $insert_field
     * @param $setSkipRows
     * @return array
     * @throws \Exception
     */
    public static function ExcelImport($filePath,$filename,array $insert_field, $setSkipRows = 0)
    {
        $xlswriter = new XlswriterService();
        return $xlswriter::import($filePath,$filename,$insert_field,$setSkipRows);
    }

    /**
     * 上传文件
     *
     * @param $parmas
     * @return void
     */
    public static function FileUpload($parmas){

        $xlswriter = new XlswriterService();
        $file = $parmas['file'];
        $fileExtra = [
            'file_size'   => $parmas['file_size'],
            'file_suffix' => $parmas['file_suffix'],
            'file_name'   => $parmas['file_name'],
        ];
        return $xlswriter::fileUpload($file, $fileExtra);
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