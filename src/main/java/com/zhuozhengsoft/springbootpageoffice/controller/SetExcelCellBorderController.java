package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.*;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.awt.*;
import java.util.Map;

@RestController
@RequestMapping(value = "/SetExcelCellBorder/")
public class SetExcelCellBorderController {

    @RequestMapping(value = "Excel", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        Workbook wb = new Workbook();
        Sheet sheet = wb.openSheet("Sheet1");
        // 设置背景
        Table backGroundTable = sheet.openTable("A1:P200");
        //设置表格边框样式
        backGroundTable.getBorder().setLineColor(Color.white);

        // 设置单元格边框样式
        Border C4Border = sheet.openTable("C4:C4").getBorder();
        C4Border.setWeight(XlBorderWeight.xlThick);
        C4Border.setLineColor(Color.yellow);
        C4Border.setBorderType(XlBorderType.xlAllEdges);

        // 设置单元格边框样式
        Border B6Border = sheet.openTable("B6:B6").getBorder();
        B6Border.setWeight(XlBorderWeight.xlHairline);
        B6Border.setLineColor(Color.magenta);
        B6Border.setLineStyle(XlBorderLineStyle.xlSlantDashDot);
        B6Border.setBorderType(XlBorderType.xlAllEdges);

        //设置表格边框样式
        Table titleTable = sheet.openTable("B4:F5");
        titleTable.getBorder().setWeight(XlBorderWeight.xlThick);
        titleTable.getBorder().setLineColor(new Color(0, 128, 128));
        titleTable.getBorder().setBorderType(XlBorderType.xlAllEdges);

        //设置表格边框样式
        Table bodyTable2 = sheet.openTable("B6:F15");
        bodyTable2.getBorder().setWeight(XlBorderWeight.xlThick);
        bodyTable2.getBorder().setLineColor(new Color(0, 128, 128));
        bodyTable2.getBorder().setBorderType(XlBorderType.xlAllEdges);

        poCtrl.setWriter(wb);

        //打开Word文档
        poCtrl.webOpen("/doc/SetExcelCellBorder/test.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("SetExcelCellBorder/Word");
        return mv;
    }


}
