package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.springbootpageoffice.util.GetDirPathUtil;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@RestController
@RequestMapping(value = "/FileMakerConvertPDFs/")
public class FileMakerConvertPDFsController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "index", method = RequestMethod.GET)
    public ModelAndView showindex(HttpServletRequest request, Map<String, Object> map) {
        map.put("url", dir + "FileMakerConvertPDFs/");
        ModelAndView mv = new ModelAndView("FileMakerConvertPDFs/index");
        return mv;
    }

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {

        String path = request.getContextPath();
        String filePath = "";
        String id = request.getParameter("id").trim();
        if ("1".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "PageOffice产品简介.doc";
        }
        if ("2".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "Pageoffice客户端安装步骤.doc";
        }
        if ("3".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "PageOffice的应用领域.doc";
        }
        if ("4".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "PageOffice产品对客户端环境要求.doc";
        }

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        filePath = filePath.replace("/", "\\");
        //打开Word文档
        poCtrl.webOpen(filePath, OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("FileMakerConvertPDFs/Word");
        return mv;
    }

    @RequestMapping(value = "Convert", method = RequestMethod.GET)
    public ModelAndView showWConvert(HttpServletRequest request, Map<String, Object> map) {
        String path = request.getContextPath();
        String filePath = "";
        String id = request.getParameter("id").trim();
        if ("1".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "PageOffice产品简介.doc";
        }
        if ("2".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "Pageoffice客户端安装步骤.doc";
        }
        if ("3".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "PageOffice的应用领域.doc";
        }
        if ("4".equals(id)) {
            filePath = dir + "FileMakerConvertPDFs\\" + "PageOffice产品对客户端环境要求.doc";
        }

        filePath = filePath.replace("/", "\\");
        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        fmCtrl.setSaveFilePage("save");
        fmCtrl.fillDocumentAsPDF(filePath, DocumentOpenType.Word, "a.pdf");
        map.put("pageoffice", fmCtrl.getHtmlCode("FileMakerCtrl1"));
        ModelAndView mv = new ModelAndView("FileMakerConvertPDFs/Convert");
        return mv;
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "FileMakerConvertPDFs/" + fs.getFileName());
        fs.close();
    }

}
