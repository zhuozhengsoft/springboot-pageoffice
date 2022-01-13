package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.Cell;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.awt.*;

import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Map;

@RestController
@RequestMapping(value = "/ExcelFill2/")
public class ExcelFill2Controller {

    @RequestMapping(value = "Excel", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        poCtrl.setCaption("简单的给Excel赋值");
        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");
        //定义Cell对象
        Cell cellB4 = sheet.openCell("B4");
        //给单元格赋值
        cellB4.setValue("1月");
        //设置字体颜色
        cellB4.setForeColor(Color.red);

        Cell cellC4 = sheet.openCell("C4");
        cellC4.setValue("300");
        cellC4.setForeColor(Color.blue);

        Cell cellD4 = sheet.openCell("D4");
        cellD4.setValue("270");
        cellD4.setForeColor(Color.orange);

        Cell cellE4 = sheet.openCell("E4");
        cellE4.setValue("270");
        cellE4.setForeColor(Color.green);

        Cell cellF4 = sheet.openCell("F4");
        DecimalFormat df = (DecimalFormat) NumberFormat.getInstance();
        cellF4.setValue(df.format(270.00 / 300 * 100) + "%");
        cellF4.setForeColor(Color.gray);

        poCtrl.setWriter(workBook);

        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏工具栏
        poCtrl.setCustomToolbar(false);
        //打开Word文档
        poCtrl.webOpen("/doc/ExcelFill2/test.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("ExcelFill2/Excel");
        return mv;
    }


}
