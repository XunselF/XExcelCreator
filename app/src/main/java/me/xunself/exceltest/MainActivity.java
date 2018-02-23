package me.xunself.exceltest;

import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;

import java.io.File;
import java.io.IOException;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import me.xunself.XExcelCreator.XExcelCreator;
import me.xunself.exceltext.R;

public class MainActivity extends AppCompatActivity implements View.OnClickListener{

    private Button createExcelButton;   //创建excel表按钮


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        initView();
    }
    @Override
    public void onClick(View view) {
        switch(view.getId()){
            case R.id.create_excel_button:
                XExcelCreator creator = new XExcelCreator("test1","1",0);
                String[] titles = {"时间","标题","内容"};
                creator.createAllTexts(titles);
                creator.setmHeight(30);
                creator.setmWidth(300);
                creator.createAllTexts(new String[]{"1","2","3"});
                String[] texts = {"2017-12-08","android","Hello World！"};
                creator.createAllTexts(texts);
                creator.writeData();
                break;
            default:
                break;
        }
    }
    /**
     * 初始化控件
     */
    private void initView(){
        createExcelButton = (Button) findViewById(R.id.create_excel_button);
        createExcelButton.setOnClickListener(this);
    }

    /**
     * 创建excel表
     */
    private void createExcel(){
        try{
            //打开文件 生成文件
            WritableWorkbook book = Workbook.createWorkbook(new File("mnt/sdcard/test.xls"));
            //生成名为“第一张工作表” 的工作表，参数0表示这是第一页。
            WritableSheet sheet = book.createSheet("第一张工作表",0);
            //设置行高，设置第一行高度为100，参数1：行数 参数2：高度

            sheet.setColumnView(0,30);
//设置列宽，设置第一列宽度50，参数1：列数 参数2：高度
            sheet.setColumnView(1,30);
//设置列宽，设置第一列宽度50，参数1：列数 参数2：高度
            sheet.setColumnView(2,30);
            //创建字体，参数1：字体样式 参数2：字号， 参数3：粗体
            WritableFont font = new WritableFont(WritableFont.ARIAL,11,WritableFont.BOLD);
            WritableCellFormat format = new WritableCellFormat(font);

            //设置对齐方式为水平居中
            format.setAlignment(Alignment.CENTRE);
            //设置对齐方式为垂直居中
            format.setVerticalAlignment(VerticalAlignment.CENTRE);

            //在Label对象的构造子指明单位格位置是第一行第一列
            //单位格内容为baby
            Label label = new Label(2,0,"baby",format);

            //将定义好的单元格添加到工作表中
            sheet.addCell(label);

            //生成一个保存数字的单位格，必须用Number的完整包路径，否则会发生语法歧义
            jxl.write.Number number = new jxl.write.Number(0,6,123.456);
            sheet.addCell(number);
            //写入数据并关闭
            book.write();
            book.close();


        }catch(WriteException e){
            e.printStackTrace();
        }
        catch(IOException e){
            e.printStackTrace();
        }
    }

}
