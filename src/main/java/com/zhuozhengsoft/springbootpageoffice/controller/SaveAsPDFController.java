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
@RequestMapping(value = "/SaveAsPDF/")
public class SaveAsPDFController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("另存为PDF文件", "SaveAsPDF()", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/SaveAsPDF/template.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("SaveAsPDF/Word");
        return mv;
    }

    @RequestMapping(value = "pdf", method = RequestMethod.GET)
    public ModelAndView showPdf(HttpServletRequest request, Map<String, Object> map) {
        PDFCtrl pdfCtrl = new PDFCtrl(request);
        pdfCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        pdfCtrl.setTheme(ThemeType.CustomStyle);//设置主题样式
        //添加自定义按钮
        pdfCtrl.addCustomToolButton("打印", "Print()", 6);
        pdfCtrl.addCustomToolButton("-", "", 0);
        pdfCtrl.addCustomToolButton("实际大小", "SetPageReal()", 16);
        pdfCtrl.addCustomToolButton("适合页面", "SetPageFit()", 17);
        pdfCtrl.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
        pdfCtrl.addCustomToolButton("-", "", 0);
        pdfCtrl.addCustomToolButton("首页", "FirstPage()", 8);
        pdfCtrl.addCustomToolButton("上一页", "PreviousPage()", 9);
        pdfCtrl.addCustomToolButton("下一页", "NextPage()", 10);
        pdfCtrl.addCustomToolButton("尾页", "LastPage()", 11);
        pdfCtrl.addCustomToolButton("-", "", 0);
        pdfCtrl.addCustomToolButton("左转", "RotateLeft()", 12);
        pdfCtrl.addCustomToolButton("右转", "RotateRight()", 13);
        pdfCtrl.addCustomToolButton("-", "", 0);
        pdfCtrl.addCustomToolButton("放大", "ZoomIn()", 14);
        pdfCtrl.addCustomToolButton("缩小", "ZoomOut()", 15);
        pdfCtrl.addCustomToolButton("-", "", 0);
        pdfCtrl.addCustomToolButton("全屏", "SwitchFullScreen()", 4);
        //设置禁止拷贝
        pdfCtrl.setAllowCopy(false);
        String fileName = request.getParameter("fileName");//定义文件名称
        pdfCtrl.webOpen("/doc/SaveAsPDF/" + fileName);

        map.put("pageoffice", pdfCtrl.getHtmlCode("PDFCtrl1"));
        ModelAndView mv = new ModelAndView("SaveAsPDF/pdf");
        return mv;
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "SaveAsPDF/" + fs.getFileName());
        fs.close();
    }

}
