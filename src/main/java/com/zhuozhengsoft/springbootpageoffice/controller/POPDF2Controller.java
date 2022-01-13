package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.PDFCtrl;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@RestController
@RequestMapping(value = "/POPDF2/")
public class POPDF2Controller {

    @RequestMapping(value = "PDF", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PDFCtrl pdfCtrl1 = new PDFCtrl(request);
        pdfCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        // Create custom toolbar
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
        pdfCtrl1.webOpen("/doc/POPDF2/test.pdf");
        map.put("pageoffice", pdfCtrl1.getHtmlCode("PDFCtrl1"));

        ModelAndView mv = new ModelAndView("POPDF2/pdf");
        return mv;
    }


}
