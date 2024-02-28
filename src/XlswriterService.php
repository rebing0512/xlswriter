<?php

namespace Jenson\Xlswriter;

use Vtiful\Kernel\Excel;

class XlswriterService
{
    public $excel;

    public $setting = array();

    public $data = array();

    /**
     * 初始化
     *
     * @param $path
     * @param $data
     * @param $setting
     * @param $type
     * @throws \Exception
     */
    public function __construct($path = '/xlswriter/viest',$type = 'write',$data = array(),$setting = array())
    {
        $config = [
            'path' => $path #xlsx文件保存路径
        ];
        $arr = [
            'read','write'
        ];
        if(!in_array($type,$arr)){
            throw new \Exception('Operate type must be read or write');
        }
        $this->excel =  new Excel($config);
        $this->data =  $data;
        $this->setting =  $setting;
    }

    /**
     * Excel导出
     *
     * @param $title
     * @return void
     * @throws \Exception
     */
    public function export($title)
    {
        $data = $this->data;
        $setting = $this->setting;
        if (count($title) == 0 || count($data) == 0) {
            throw new \Exception('The title is not empty!');
        }
        $defaultSetting = [
            'fileName'=>'excel_export_'.date('Y-m-d').'.xlsx',

            'hasSheetTitle'=>true, # 是否有表格的表头

            'hasSerialNumber'=>true, # 是否有编号
            'serialNumberTitle'=>"编号", # 是否有编号

            'hasFreezePane'=>true, # 是否进行冻结分割
            'freezePane'=>[1,0], # 第几行，第几列

            'sheetName'=>'工作表sheet标题',
            'titleName'=>'工作表#sheetName#title标题',

            'defaultRowWidth'=>10,  # 默认行宽
            'defaultRowHeight'=>30, # 默认行高

            'headerStyleSize'=>18,  # 表头的文字大小，加粗
            'headerRowHeight'=>48,  # 表头的的行高

            'titleStyleSize'=>11,  # 标题内容的字体大小，加粗默认11
            'titleRowHeight'=>30,  # 表头的的行高
            //'date'=>'', # 日期信息

            'hasTotalInfo'=>true, #  是否包含合计信息
            'totalInfo'=>'共计#COUNT#条', #  合计信息内容
        ];
        $setting = array_merge($defaultSetting,$setting);
        $excel = $this->excel;
        # 冻结从几行、第几列开始
        $hasFreezePane = $setting['hasFreezePane'];
        if($hasFreezePane){
            $excel->freezePanes($setting['freezePane'][0],$setting['freezePane'][1]);
        }
        $sheetName =$setting['sheetName'];
        $fileName =$setting['fileName'];
        $filePath = $excel->fileName($fileName, $sheetName)
            ->header(['Item', 'Cost'])
            ->data([
                ['Rent', 1000],
                ['Gas',  100],
                ['Food', 300],
                ['Gym',  50],
            ])
            ->output();
    }

    /**
     * 数字续号对应的字母
     *
     * @var string[]
     */
    private static $keyArr = [
        1 => 'A',
        2 => 'B',
        3 => 'C',
        4 => 'D',
        5 => 'E',
        6 => 'F',
        7 => 'G',
        8 => 'H',
        9 => 'I',
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

}