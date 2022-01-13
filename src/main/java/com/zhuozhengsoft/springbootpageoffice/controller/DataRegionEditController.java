package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.*;
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

@RestController
@RequestMapping(value = "/DataRegionEdit/")
public class DataRegionEditController {
    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        WordDocument doc = new WordDocument();
        doc.getTemplate().defineDataRegion("Name", "[ 姓名 ]");
        doc.getTemplate().defineDataRegion("Address", "[ 地址 ]");
        doc.getTemplate().defineDataRegion("Tel", "[ 电话 ]");
        doc.getTemplate().defineDataRegion("Phone", "[ 手机 ]");
        doc.getTemplate().defineDataRegion("Sex", "[ 性别 ]");
        doc.getTemplate().defineDataRegion("Age", "[ 年龄 ]");
        doc.getTemplate().defineDataRegion("Email", "[ 邮箱 ]");
        doc.getTemplate().defineDataRegion("QQNo", "[ QQ号 ]");
        doc.getTemplate().defineDataRegion("MSNNo", "[ MSN号 ]");

        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("定义数据区域", "ShowDefineDataRegions()", 3);

        poCtrl.setTheme(ThemeType.Office2007);
        poCtrl.setBorderStyle(BorderStyleType.BorderThin);
        poCtrl.setWriter(doc);
        poCtrl.setSaveFilePage("save");
        //打开Word文档
        poCtrl.webOpen("/doc/DataRegionEdit/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DataRegionEdit/Word");
        return mv;
    }

    @RequestMapping(value = "DataRegionDlg", method = RequestMethod.GET)
    public ModelAndView dataRegionDlg(HttpServletRequest request, Map<String, Object> map) {

        return new ModelAndView("DataRegionEdit/DataRegionDlg");
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "DataRegionEdit/" + fs.getFileName());
        fs.close();
    }

}
