package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
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
import java.util.Map;

@RestController
@RequestMapping(value = "/SetXlsTableByUser/")
public class SetXlsTableByUserController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "index", method = RequestMethod.GET)
    public ModelAndView showindex(HttpServletRequest request, Map<String, Object> map) {
        ModelAndView mv = new ModelAndView("SetXlsTableByUser/index");
        return mv;
    }

    @RequestMapping(value = "Excel", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        String userName = request.getParameter("userName");

        Workbook wb = new Workbook();
        Sheet sheet = wb.openSheet("Sheet1");
        Table tableA = sheet.openTable("C4:D6");
        Table tableB = sheet.openTable("C7:D9");

        tableA.setSubmitName("tableA");
        tableB.setSubmitName("tableB");

        //根据登录用户名设置数据区域可编辑性
        String strInfo = "";

        //A部门经理登录后
        if (userName.equals("zhangsan")) {
            strInfo = "A部门经理，所以只能编辑A部门的产品数据";
            tableA.setReadOnly(false);
            tableB.setReadOnly(true);
        }
        //B部门经理登录后
        else {
            strInfo = "B部门经理，所以只能编辑B部门的产品数据";
            tableA.setReadOnly(true);
            tableB.setReadOnly(false);
        }

        poCtrl.setWriter(wb);
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/SetXlsTableByUser/test.xls", OpenModeType.xlsSubmitForm, userName);
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("strInfo", strInfo);
        ModelAndView mv = new ModelAndView("SetXlsTableByUser/Excel");
        return mv;
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "SetXlsTableByUser/" + fs.getFileName());
        fs.close();
    }

}
