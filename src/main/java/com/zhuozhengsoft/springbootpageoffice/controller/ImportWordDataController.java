package com.zhuozhengsoft.springbootpageoffice.controller;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import com.zhuozhengsoft.springbootpageoffice.util.GetDirPathUtil;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.util.Map;

import static java.lang.System.out;

@RestController
@RequestMapping(value = "/ImportWordData/")
public class ImportWordDataController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "/Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        poCtrl.addCustomToolButton("导入文件", "importData()", 15);
        poCtrl.addCustomToolButton("提交数据", "submitData()", 1);

        WordDocument doc = new WordDocument();
        poCtrl.setWriter(doc);
        //设置保存页面
        poCtrl.setSaveDataPage("save");//设置处理文件保存的请求方法
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("ImportWordData/Word");
        return mv;
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        String ErrorMsg = "";
        String BaseUrl = "";
        //-----------  PageOffice 服务器端编程开始  -------------------//
        com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);
        String sName = doc.openDataRegion("PO_name").getValue();
        String sDept = doc.openDataRegion("PO_dept").getValue();
        String sCause = doc.openDataRegion("PO_cause").getValue();
        String sNum = doc.openDataRegion("PO_num").getValue();
        String sDate = doc.openDataRegion("PO_date").getValue();

        if (sName.equals("")) {
            ErrorMsg = ErrorMsg + "<li>申请人</li>";
        }
        if (sDept.equals("")) {
            ErrorMsg = ErrorMsg + "<li>部门名称</li>";
        }
        if (sCause.equals("")) {
            ErrorMsg = ErrorMsg + "<li>请假原因</li>";
        }
        if (sDate.equals("")) {
            ErrorMsg = ErrorMsg + "<li>日期</li>";
        }
        try {
            if (sNum != "") {
                if (Integer.parseInt(sNum) < 0) {
                    ErrorMsg = ErrorMsg + "<li>请假天数不能是负数</li>";
                }
            } else {
                ErrorMsg = ErrorMsg + "<li>请假天数</li>";
            }
        } catch (Exception Ex) {
            ErrorMsg = ErrorMsg + "<li><font color=red>注意：</font>请假天数必须是数字</li>";
        }

        if (ErrorMsg == "") {
            // 您可以在此编程，保存这些数据到数据库中。
            out.println("提交的数据为：<br/>");
            out.println("姓名：" + sName + "<br/>");
            out.println("部门：" + sDept + "<br/>");
            out.println("原因：" + sCause + "<br/>");
            out.println("天数：" + sNum + "<br/>");
            out.println("日期：" + sDate + "<br/>");
            doc.showPage(578, 380);
        } else {
            ErrorMsg = "<div style='color:#FF0000;'>请修改以下信息：</div> "
                    + ErrorMsg;
            doc.showPage(578, 380);
        }
        doc.close();
    }

}
