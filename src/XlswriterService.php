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
            'fileName'=>'excel_export_'.date('Y-m-d H:i:s').'.xlsx',

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

}