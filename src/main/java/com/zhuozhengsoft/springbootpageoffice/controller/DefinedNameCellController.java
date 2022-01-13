package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@RestController
@RequestMapping(value = "/DefinedNameCell/")
public class DefinedNameCellController {

    @RequestMapping(value = "Excel", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        poCtrl.setCaption("简单的给Excel赋值");
        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");
        sheet.openCellByDefinedName("testA1").setValue("Tom");
        sheet.openCellByDefinedName("testB1").setValue("John");

        poCtrl.setWriter(workBook);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        //设置保存页面
        poCtrl.setSaveDataPage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/DefinedNameCell/test.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DefinedNameCell/Excel");
        return mv;
    }


    @RequestMapping("save")
    public ModelAndView save(HttpServletRequest request, HttpServletResponse response, Map<String, Object> map) {
        com.zhuozhengsoft.pageoffice.excelreader.Workbook workBook = new com.zhuozhengsoft.pageoffice.excelreader.Workbook(request, response);
        com.zhuozhengsoft.pageoffice.excelreader.Sheet sheet = workBook.openSheet("Sheet1");
        String content = "";
        content += "testA1：" + sheet.openCellByDefinedName("testA1").getValue() + "<br/>";
        content += "testB1：" + sheet.openCellByDefinedName("testB1").getValue() + "<br/>";

        workBook.showPage(500, 400);
        workBook.close();

        map.put("content", content);
        ModelAndView mv = new ModelAndView("DefinedNameCell/save");
        return mv;
    }

}
