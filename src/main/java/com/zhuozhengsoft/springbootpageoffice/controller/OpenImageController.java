package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.PDFCtrl;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@RestController
@RequestMapping(value = "/OpenImage/")
public class OpenImageController {
    @RequestMapping(value = "PDF", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PDFCtrl poCtrl = new PDFCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        // Create custom toolbar
        poCtrl.addCustomToolButton("打印", "Print()", 6);
        poCtrl.addCustomToolButton("-", "", 0);
        poCtrl.addCustomToolButton("实际大小", "SetPageReal()", 16);
        poCtrl.addCustomToolButton("适合页面", "SetPageFit()", 17);
        poCtrl.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
        poCtrl.addCustomToolButton("-", "", 0);
        poCtrl.addCustomToolButton("左转", "RotateLeft()", 12);
        poCtrl.addCustomToolButton("右转", "RotateRight()", 13);
        poCtrl.addCustomToolButton("-", "", 0);
        poCtrl.addCustomToolButton("放大", "ZoomIn()", 14);
        poCtrl.addCustomToolButton("缩小", "ZoomOut()", 15);
        poCtrl.addCustomToolButton("-", "", 0);
        poCtrl.addCustomToolButton("全屏", "SwitchFullScreen()", 4);
        poCtrl.webOpen("/doc/OpenImage/test.jpg");

        map.put("pageoffice", poCtrl.getHtmlCode("PDFCtrl1"));
        ModelAndView mv = new ModelAndView("OpenImage/pdf");
        return mv;
    }

}
