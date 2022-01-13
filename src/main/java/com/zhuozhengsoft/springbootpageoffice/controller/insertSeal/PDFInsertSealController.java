package com.zhuozhengsoft.springbootpageoffice.controller.insertSeal;

import com.zhuozhengsoft.pageoffice.*;
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
@RequestMapping(value = "/InsertSeal/PDF/")
public class PDFInsertSealController {
    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "AddSeal/PDF1", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PDFCtrl pdfCtrl1 = new PDFCtrl(request);
        pdfCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //设置保存页面
        pdfCtrl1.setSaveFilePage("save");

        // Create custom toolbar
        pdfCtrl1.addCustomToolButton("保存", "Save()", 1);
        pdfCtrl1.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        pdfCtrl1.addCustomToolButton("打印", "PrintFile()", 6);
        pdfCtrl1.addCustomToolButton("隐藏/显示书签", "SetBookmarks()", 0);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("实际大小", "SetPageReal()", 16);
        pdfCtrl1.addCustomToolButton("适合页面", "SetPageFit()", 17);
        pdfCtrl1.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("首页", "FirstPage()", 8);
        pdfCtrl1.addCustomToolButton("上一页", "PreviousPage()", 9);
        pdfCtrl1.addCustomToolButton("下一页", "NextPage()", 10);
        pdfCtrl1.addCustomToolButton("尾页", "LastPage()", 11);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("向左旋转90度", "SetRotateLeft()", 12);
        pdfCtrl1.addCustomToolButton("向右旋转90度", "SetRotateRight()", 13);
        pdfCtrl1.webOpen("/doc/InsertSeal/PDF/AddSeal/test1.pdf");

        map.put("pageoffice", pdfCtrl1.getHtmlCode("PDFCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/PDF/AddSeal/PDF1");
        return mv;
    }


    @RequestMapping(value = "AddSign/PDF1", method = RequestMethod.GET)
    public ModelAndView showWord11(HttpServletRequest request, Map<String, Object> map) {
        PDFCtrl pdfCtrl1 = new PDFCtrl(request);
        pdfCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须

        //设置保存页面
        pdfCtrl1.setSaveFilePage("save");

        // Create custom toolbar
        pdfCtrl1.addCustomToolButton("保存", "Save()", 1);
        pdfCtrl1.addCustomToolButton("签字", "AddHandSign()", 3);
        pdfCtrl1.addCustomToolButton("打印", "PrintFile()", 6);
        pdfCtrl1.addCustomToolButton("隐藏/显示书签", "SetBookmarks()", 0);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("实际大小", "SetPageReal()", 16);
        pdfCtrl1.addCustomToolButton("适合页面", "SetPageFit()", 17);
        pdfCtrl1.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("首页", "FirstPage()", 8);
        pdfCtrl1.addCustomToolButton("上一页", "PreviousPage()", 9);
        pdfCtrl1.addCustomToolButton("下一页", "NextPage()", 10);
        pdfCtrl1.addCustomToolButton("尾页", "LastPage()", 11);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("向左旋转90度", "SetRotateLeft()", 12);
        pdfCtrl1.addCustomToolButton("向右旋转90度", "SetRotateRight()", 13);
        pdfCtrl1.webOpen("/doc/InsertSeal/PDF/AddSign/test1.pdf");

        map.put("pageoffice", pdfCtrl1.getHtmlCode("PDFCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/PDF/AddSign/PDF1");
        return mv;
    }


    @RequestMapping("AddSeal/save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "InsertSeal/PDF/AddSeal/" + fs.getFileName());
        fs.close();
    }

    @RequestMapping("AddSign/save")
    public void save2(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "InsertSeal/PDF/AddSign/" + fs.getFileName());
        fs.close();
    }

}
