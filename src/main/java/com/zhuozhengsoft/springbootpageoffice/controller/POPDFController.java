package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.PDFCtrl;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@RestController
@RequestMapping(value = "/POPDF/")
public class POPDFController {

    @RequestMapping(value = "pdf", method = RequestMethod.GET)
    public ModelAndView showPdf(HttpServletRequest request, Map<String, Object> map) {
        PDFCtrl pdfCtrl = new PDFCtrl(request);
        pdfCtrl.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        // Create custom toolbar
        pdfCtrl.addCustomToolButton("打印", "PrintFile()", 6);
        pdfCtrl.addCustomToolButton("隐藏/显示书签", "SetBookmarks()", 0);
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
        pdfCtrl.addCustomToolButton("向左旋转90度", "SetRotateLeft()", 12);
        pdfCtrl.addCustomToolButton("向右旋转90度", "SetRotateRight()", 13);
        pdfCtrl.webOpen("/doc/POPDF/test.pdf");

        map.put("pageoffice", pdfCtrl.getHtmlCode("PDFCtrl1"));
        ModelAndView mv = new ModelAndView("POPDF/pdf");
        return mv;
    }


}
