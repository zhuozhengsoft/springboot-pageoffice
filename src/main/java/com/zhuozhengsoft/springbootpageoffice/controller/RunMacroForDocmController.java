package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

/**
 * @Author: dong
 * @Date: 2020/11/6 9:36
 * @Version 1.0
 */

@RestController
@RequestMapping(value = "/RunMacroForDocm")
public class RunMacroForDocmController {
    @RequestMapping(value = "/Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl1 = new PageOfficeCtrl(request);
        poCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //隐藏菜单栏
        poCtrl1.setMenubar(false);
        //隐藏自定义工具栏
        poCtrl1.setCustomToolbar(false);
        //打开文件
        poCtrl1.webOpen("/doc/RunMacroForDocm/test.docm", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl1.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("RunMacroForDocm/Word");
        return mv;
    }
}
