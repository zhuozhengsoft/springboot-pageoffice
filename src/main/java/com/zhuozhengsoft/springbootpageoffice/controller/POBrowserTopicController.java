package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.springbootpageoffice.util.GetDirPathUtil;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.FileNotFoundException;
import java.util.Map;

@RestController
@RequestMapping(value = "/POBrowserTopic/")
public class POBrowserTopicController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "index", method = RequestMethod.GET)
    public ModelAndView showIndex(HttpServletRequest request, Map<String, Object> map, HttpSession session) {
        String userName = "张三";
        session.setAttribute("userName", userName);
        ModelAndView mv = new ModelAndView("POBrowserTopic/index");
        return mv;
    }


    @RequestMapping(value = "Word1", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/POBrowserTopic/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("POBrowserTopic/Word1");
        return mv;
    }


    @RequestMapping(value = "Word2", method = RequestMethod.GET)
    public ModelAndView showWord2(HttpServletRequest request, Map<String, Object> map, HttpSession session) {
        //获取index页面传递过来参数的值
        String userName = (String) session.getAttribute("userName");
        //获取index用？传递过来的id的值
        String id = request.getParameter("id");
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/POBrowserTopic/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("userName", userName);
        map.put("id", id);

        ModelAndView mv = new ModelAndView("POBrowserTopic/Word2");
        return mv;
    }

    @RequestMapping(value = "Word3", method = RequestMethod.GET)
    public ModelAndView showWord3(HttpServletRequest request, Map<String, Object> map, HttpSession session) {
        String txt = (String) session.getAttribute("txt");
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存并关闭", "Save", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/POBrowserTopic/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("txt", txt);

        ModelAndView mv = new ModelAndView("POBrowserTopic/Word3");
        return mv;
    }


    @RequestMapping("Result2")
    @ResponseBody
    public String Result2(HttpServletRequest request, HttpServletResponse response, HttpSession session) {
        String paramValue = request.getParameter("param");
        session.setAttribute("txt", paramValue);
        return "ok";
    }

    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "POBrowserTopic/" + fs.getFileName());
        fs.close();
    }

}
