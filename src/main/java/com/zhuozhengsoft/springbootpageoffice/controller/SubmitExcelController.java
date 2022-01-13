package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Table;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;
import com.zhuozhengsoft.springbootpageoffice.util.GetDirPathUtil;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Map;

@RestController
@RequestMapping(value = "/SubmitExcel/")
public class SubmitExcelController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "Excel", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");
        //定义table对象，设置table对象的设置范围
        Table table = sheet.openTable("B4:F13");
        //设置table对象的提交名称，以便保存页面获取提交的数据
        table.setSubmitName("Info");

        poCtrl.setWriter(workBook);
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveDataPage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/SubmitExcel/test.xls", OpenModeType.xlsSubmitForm, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("SubmitExcel/Word");
        return mv;
    }


    @RequestMapping("save")
    public ModelAndView save(HttpServletRequest request, HttpServletResponse response, Map<String, Object> map) {
        response.setCharacterEncoding("utf-8");//解决返回的数据中文乱码问题
        com.zhuozhengsoft.pageoffice.excelreader.Workbook workBook = new com.zhuozhengsoft.pageoffice.excelreader.Workbook(request, response);
        com.zhuozhengsoft.pageoffice.excelreader.Sheet sheet = workBook.openSheet("Sheet1");
        com.zhuozhengsoft.pageoffice.excelreader.Table table = sheet.openTable("Info");
        String content = "";
        int result = 0;
        while (!table.getEOF()) {
            //获取提交的数值
            if (!table.getDataFields().getIsEmpty()) {
                content += "<br/>月份名称："
                        + table.getDataFields().get(0).getText();
                content += "<br/>计划完成量："
                        + table.getDataFields().get(1).getText();
                content += "<br/>实际完成量："
                        + table.getDataFields().get(2).getText();
                content += "<br/>累计完成量："
                        + table.getDataFields().get(3).getText();
                //out.print(table.getDataFields().get(2).getText()+"      mmmmmmmmmmmmm          "+table.getDataFields().get(1).getText());
                if (table.getDataFields().get(2).getText().equals(null)
                        || table.getDataFields().get(2).getText().trim().length() == 0
                ) {
                    content += "<br/>完成率：0%";
                } else {
                    float f = Float.parseFloat(table.getDataFields().get(2)
                            .getText());
                    f = f / Float.parseFloat(table.getDataFields().get(1).getText());
                    DecimalFormat df = (DecimalFormat) NumberFormat.getInstance();
                    content += "<br/>完成率：" + df.format(f * 100) + "%";
                }
                content += "<br/>*********************************************";
            }
            //循环进入下一行
            table.nextRow();
        }
        table.close();

        workBook.showPage(500, 400);
        workBook.close();
        map.put("content", content);
        ModelAndView mv = new ModelAndView("SubmitExcel/save");
        return mv;
    }

}
