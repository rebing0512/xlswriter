<?php

namespace Jenson\Xlswriter\Service;

use Vtiful\Kernel\Excel;
use Vtiful\Kernel\Format;

class XlswriterService
{
    public $excel;

    /**
     * 初始化
     *
     * @param string $path
     * @throws \Exception
     */
    public function __construct(string $path = '')
    {
        $config = [
            'path' => $path #文件保存路径
        ];
        $this->excel  = new Excel($config);
    }

    /**
     *  Excel导出
     *
     * @param $title
     * @param $data
     * @param array $setting
     * @return string|void
     * @throws \Exception
     */
    public function export($title, $data, array $setting = array())
    {
        if (count($title) == 0 || count($data) == 0) {
            throw new \Exception('暂无可导出数据！');
        }
        $defaultSetting = [
            'fileName'=>'jenson_excel_export_'.date('Y-m-d').'.csv',

            'hasSheetTitle'=>true,       #是否有表格的表头
            'hasSerialNumber'=>true,     #是否有编号
            'serialNumberTitle'=>"编号",  #是否有编号

            'hasFreezePane'=>true,       #是否进行冻结分割
            'freezePane'=>[1,0],         #第几行，第几列 默认首行

            'sheetName'=>'工作表sheet标题',
            'titleName'=>'工作表#sheetName#title标题',

            'defaultRowWidth'=>15,       #默认行宽
            'defaultRowHeight'=>25,      #默认行高

            'headerStyleSize'=>18,       #表头的文字大小
            'headerRowHeight'=>48,       #表头的的行高
            'headerFont'=>'Calibri',     #表头的的字体

            'titleStyleSize'=>14,        #标题内容的字体大小
            'titleRowHeight'=>30,        #标题的的行高
            'titleFont'=>'Calibri',      #标题的的字体

            'contentStyleSize'=>12,      #内容内容的字体大小
            'contentRowHeight'=>15,      #内容的的行高
            'contentFont'=>'Calibri',    #内容的的字体

            'hasTotalInfo'=>true,        #是否包含合计信息
            'totalInfo'=>'共计#COUNT#条', #合计信息内容
        ];
        $setting = array_merge($defaultSetting,$setting);
        $excel = $this->excel;
        $sheetName  = $setting['sheetName'];
        $fileName   = $setting['fileName'];
        $excel->fileName($fileName, $sheetName);
        # 冻结从几行、第几列开始
        $hasFreezePane = $setting['hasFreezePane'];
        if($hasFreezePane){
            $excel->freezePanes($setting['freezePane'][0],$setting['freezePane'][1]);
        }
        $fileHandle = $excel->getHandle();
        $format1    = new \Vtiful\Kernel\Format($fileHandle);
        $format2    = new \Vtiful\Kernel\Format($fileHandle);
        #title style
        $titleStyle = $format1->fontSize($setting['headerStyleSize'])
            ->bold()
            ->font($setting['titleFont'])
            ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
            ->toResource();
        #global style
        $globalStyle = $format2->fontSize($setting['contentStyleSize'])
            ->font($setting['contentFont'])
            ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
            ->border(Format::BORDER_THIN)
            ->toResource();
        $header = $title;
        $headerLen = count($header)-1;
        #header
        $list = $data;
        array_unshift($list, $header);
        $title = array_fill(1, $headerLen, '');
        $titleName   = $setting['titleName'];
        $title[0] = $titleName;
        array_unshift($list, $title);
        $end = static::getStr($headerLen);//strtoupper(chr(65 + $headerLen));
        #column style 列宽
        $excel->setColumn("A:{$end}", $setting['defaultRowWidth'], $globalStyle);
        #title
        $excel->MergeCells("A1:{$end}1", $titleName)->setRow("A1", $setting['defaultRowHeight'], $titleStyle); # setRow 行高
        #数据
        $filePath = $excel->data($list)->output();
        #获取要下载的文件名
        $file = $filePath;
        try {
            #检查文件是否存在
            if (file_exists($filePath)) {
                #设置下载头信息
                header('Content-Description: File Transfer');
                header('Content-Type: application/octet-stream');
                header('Content-Disposition: attachment; filename="' . basename($file) . '"');
                header('Content-Transfer-Encoding: binary');
                header('Expires: 0');
                header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
                header('Pragma: public');
                header('Content-Length: ' . filesize($file));
                #读取文件并将其发送到用户
                ob_clean();
                flush();
                readfile($file);
                #删除文件（节省空间）
                unlink($filePath);
                exit;
            } else {
                echo '文件不存在';
            }
            return 'success';
        } catch (\Vtiful\Kernel\Exception $e) {
            return $e->getMessage();
        }
    }

    /**
     * @param string $filePath
     * @param string $filename
     * @param array $insert_field
     * @param int $setSkipRows
     * @return array
     * @throws \Exception
     */
    public function import(string $filePath, string $filename, array $insert_field, int $setSkipRows = 0){
        $xlsObj  = $this->excel;
        #excel文件
//        $file = "users_data.xlsx";
        //实例化reader
        $filePath .= $filename;
        $ext = pathinfo($filePath, PATHINFO_EXTENSION);
        if (!in_array($ext, ['csv', 'xls', 'xlsx'])) {
            throw new \Exception('未知文件格式！');
        }
        //打开xls文件
        $sheetList = $xlsObj->openFile($filename)->sheetList();
        //循环读取工作表
        $insert = array();
        foreach ($sheetList as $sheetName) {
            #读取工作表内容
            $xlsObj->openSheet($sheetName);
            if($setSkipRows > 0) {
                #跳过的行数
                $xlsObj->setSkipRows($setSkipRows);
            }
            $data = $xlsObj->getSheetData();
            foreach ($data as $key => $value) {
                $insert[] = [
                    $insert_field[$key] => $value[$key],
                    'add_time' => time(),
                ];
                if(in_array('upd_time',$insert_field)){
                    $insert[] = [
                        'upd_time' => time(),
                    ];
                }
            }
//            #游标模式读取数据
//            $name_exist = [];
//            while (($row = $xlsObj->nextRow()) !== NULL) {
//
//            }
        }
        return $insert;
    }
    /**
     * @var string[]
     */
    private static $keyArr = [
        1  => 'A',
        2  => 'B',
        3  => 'C',
        4  => 'D',
        5  => 'E',
        6  => 'F',
        7  => 'G',
        8  => 'H',
        9  => 'I',
        10 => 'J',
        11 => 'K',
        12 => 'L',
        13 => 'M',
        14 => 'N',
        15 => 'O',
        16 => 'P',
        17 => 'Q',
        18 => 'R',
        19 => 'S',
        20 => 'T',
        21 => 'U',
        22 => 'V',
        23 => 'W',
        24 => 'X',
        25 => 'Y',
        26 => 'Z',
    ];

    /**
     * 字母对应的数字续号
     *
     * @var int[]
     */
    private static $keyArrToNum = [
        'A' => 1,
        'B' => 2,
        'C' => 3,
        'D' => 4,
        'E' => 5,
        'F' => 6,
        'G' => 7,
        'H' => 8,
        'I' => 9,
        'J' => 10,
        'K' => 11,
        'L' => 12,
        'M' => 13,
        'N' => 14,
        'O' => 15,
        'P' => 16,
        'Q' => 17,
        'R' => 18,
        'S' => 19,
        'T' => 20,
        'U' => 21,
        'V' => 22,
        'W' => 23,
        'X' => 24,
        'Y' => 25,
        'Z' => 26,
    ];

    /**
     * 获取列的字母编号
     *
     * @param $str
     * @return float|int
     */
    public function getStrToNum($str){
        $str = strtoupper($str);
        $len = strlen($str);
        $num = 0;
        for($i=0;$i<$len;$i++){
            $num += static::$keyArrToNum[substr($str,$i,1)]*pow(26,$len-1-$i);
        }
        return $num;
    }

    /**
     * 获取列的字母编号
     *
     * @param $num
     * @return false|string
     */
    public function getStr($num){
        #如果不是整数，或小于0
        if(!is_int($num) || $num<=0){
            return false;
        }else if($num<=26){
            return self::$keyArr[$num];
        }else{
            #取余数
            $num2 = $num%26;
            if($num2==0){
                $num2 = 26;
                #取整数
                $num = intval(floor(($num-26)/26));
                return self::getStr($num).self::getStr($num2);
            }else{
                #取整数
                $num = intval(floor($num/26));
                return self::getStr($num).self::getStr($num2);
            }
        }
    }

    /**
     * 文件上传
     *
     * @param $file
     * @param $fileExtra
     * @param string $save_path
     * @return array
     */
    public function fileUpload($file, $fileExtra, string $save_path='static/upload/files'): array
    {
        if (!$file){
            return [
                'code' => 0,
                'msg' => '文件不存在'
            ];
        }
        #如果之前的文件存在
        $storage_path = $save_path;//'static/upload/files';
        $content = file_get_contents($file['tmp_name']);
        $file_path = $storage_path.'/'.$fileExtra['file_name'];
        if($fileExtra['file_index']==1 && is_file($file_path)){
            unlink($file_path);
        }
        $dir = dirname($file_path);
        if (!is_dir($dir))
        {
            mkdir($dir, 0755, true);
        }
        #写入方式打开，将文件指针指向文件末尾。如果文件不存在则尝试创建之。
        $fp = fopen($file_path, 'a');
        flock($fp, LOCK_EX);
        fwrite($fp, $content);
        flock($fp, LOCK_UN);
        #关闭
        fclose($fp);
        if (!is_file($file_path)){
            return [
                'code' => 0,
                'msg' => '文件上传失败',
                'dev' => is_file($file_path)
            ];
        }
        if($fileExtra['file_index'] < $fileExtra['file_total']){
            return [
                'code' => 1,
                'upload' => 'success'
            ];
        }else{
            $url_path = '/'.$fileExtra['file_name'].$fileExtra['file_suffix'];
            #todo:数据入库
            return [
                'code' => 1,
                'path' => $url_path,
            ];
        }
    }
}