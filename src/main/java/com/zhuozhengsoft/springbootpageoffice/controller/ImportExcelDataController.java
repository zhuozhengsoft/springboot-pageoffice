package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelreader.Table;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;
import com.zhuozhengsoft.springbootpageoffice.util.GetDirPathUtil;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Map;

@RestController
@RequestMapping(value = "/ImportExcelData/")
public class ImportExcelDataController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "Excel", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        poCtrl.addCustomToolButton("导入文件", "importData()", 16);
        poCtrl.addCustomToolButton("提交数据", "submitData()", 1);
        Workbook wb = new Workbook();
        Sheet sheet = wb.openSheet("Sheet1");
        poCtrl.setWriter(wb);
        //设置保存页面
        poCtrl.setSaveDataPage("save");//设置处理文件保存的请求方法

        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("ImportExcelData/Word");
        return mv;
    }


    @RequestMapping("save")
    public ModelAndView save(HttpServletRequest request, HttpServletResponse response, Map<String, Object> map) {
        com.zhuozhengsoft.pageoffice.excelreader.Workbook workBook = new com.zhuozhengsoft.pageoffice.excelreader.Workbook(request, response);
        com.zhuozhengsoft.pageoffice.excelreader.Sheet sheet = workBook.openSheet("Sheet1");
        Table table = sheet.openTable("B4:F13");
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
                content += "</br>";
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
