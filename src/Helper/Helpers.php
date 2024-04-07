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
     * @param $filePath
     * @param $filename
     * @param array $insert_field
     * @param int $setSkipRows
     * @return array
     * @throws \Exception
     */
    public static function ExcelImport($filePath, $filename, array $insert_field, int $setSkipRows = 0): array
    {
        $xlswriter = new XlswriterService();
        return $xlswriter->import($filePath,$filename,$insert_field,$setSkipRows);
    }

    /**
     * 上传文件
     *
     * @param $parmas
     * @return array
     * @throws \Exception
     */
    public static function FileUpload($parmas): array
    {

        $xlswriter = new XlswriterService();
        $file = $parmas['file'];
        $fileExtra = [
            'file_index'  => $file['index']??1,
            'file_size'   => $file['size'],
            'file_suffix' => $file['type'],
            'file_name'   => $file['name'],
            'file_total'  => $file['file_total']??1,
        ];
        try {
            return $xlswriter->fileUpload($file, $fileExtra);
        }Catch(\Exception $e){
            throw new \Exception($e->getMessage());
        }
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