# Jenson-xlswriter
php xlswriter插件
需要安装php_xlswriter扩展

## 安装  
```bash
composer require jenson0512/xlswriter
```
## 使用
* 1、Excel导出
    ```
    $excel = new \Jenson\Xlswriter\Helper();
    $excel->ExcelExport($title,$data,$setting = array());
    ```
* 2、Excel导入
    ```
    $excel = new \Jenson\Xlswriter\Helper();
    $excel->ExcelImport($filePath,$filename,array $insert_field, $setSkipRows = 0);
    ```
* 3、文件上传
    ```
    $excel = new \Jenson\Xlswriter\Helper();
    $excel->FileDownload($file_url);
    ```
  * $file_url：文件地址
