# Jenson-xlswriter
* php xlswriter插件（Excel大量数据导入导出插件）
* 需要安装php_xlswriter扩展
* php版本需要在7.2及以上

## 安装  
```bash
composer require jenson0512/xlswriter
```
## 使用
* 1、Excel导出
    ```
    $excel = new \Jenson\Xlswriter\Helper();
    $excel->ExcelExport($path,$title,$data,$setting = array());
    ```
* 2、Excel导入
    ```
    $excel = new \Jenson\Xlswriter\Helper();
    $excel->ExcelImport($table,$filePath,$filename,array $insert_field, $setSkipRows = 0);
    ```
* 3、文件上传
    ```
    $excel = new \Jenson\Xlswriter\Helper();
    $excel->FileUpload($parmas);
    ```
  * $parmas：必要参数
* 4、文件下载
    ```
    $excel = new \Jenson\Xlswriter\Helper();
    $excel->FileDownload($file_url);
    ```
  * $file_url：文件地址
