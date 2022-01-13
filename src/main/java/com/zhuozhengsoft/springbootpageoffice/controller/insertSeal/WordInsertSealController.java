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
@RequestMapping(value = "/InsertSeal/Word/")
public class WordInsertSealController {
    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";


    @RequestMapping(value = "BatchAddSeal/Default", method = RequestMethod.GET)
    public ModelAndView showindex(HttpServletRequest request, Map<String, Object> map) {
        map.put("url", dir + "InsertSeal/Word/BatchAddSeal/");
        ModelAndView mv = new ModelAndView("InsertSeal/word/BatchAddSeal/Default");
        return mv;
    }

    @RequestMapping(value = "BatchAddSeal/Edit", method = RequestMethod.GET)
    public ModelAndView showBatchAddSealEdit(HttpServletRequest request, Map<String, Object> map) {
        String filePath = "";
        String id = request.getParameter("id").trim();
        if ("1".equals(id)) {
            filePath = "test1.doc";
        }
        if ("2".equals(id)) {
            filePath = "test2.doc";
        }
        if ("3".equals(id)) {
            filePath = "test3.doc";
        }
        if ("4".equals(id)) {
            filePath = "test4.doc";
        }

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/BatchAddSeal/" + filePath, OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/BatchAddSeal/Edit");
        return mv;
    }

    @RequestMapping(value = "BatchAddSeal/AddSeal", method = RequestMethod.GET)
    public ModelAndView showBatchAddSealAddSeal(HttpServletRequest request, Map<String, Object> map) {

        String path = request.getContextPath();
        String filePath = "";
        String id = request.getParameter("id").trim();
        if ("1".equals(id)) {
            filePath = "test1.doc";
        }
        if ("2".equals(id)) {
            filePath = "test2.doc";
        }
        if ("3".equals(id)) {
            filePath = "test3.doc";
        }
        if ("4".equals(id)) {
            filePath = "test4.doc";
        }

        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        fmCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        //设置保存页面
        fmCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        fmCtrl.fillDocument("/doc/InsertSeal/Word/BatchAddSeal/" + filePath, DocumentOpenType.Word);

        map.put("pageoffice", fmCtrl.getHtmlCode("FileMakerCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/BatchAddSeal/AddSeal");
        return mv;
    }


    @RequestMapping("BatchAddSeal/save")
    public void save3(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "InsertSeal/Word/BatchAddSeal/" + fs.getFileName());
        fs.close();
    }


    @RequestMapping(value = "AddSeal/Word1", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

//添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal/test1.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal/Word1");
        return mv;
    }


    @RequestMapping(value = "AddSeal/Word2", method = RequestMethod.GET)
    public ModelAndView showWord2(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        //设置保存页面
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal/test2.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal/Word2");
        return mv;
    }

    @RequestMapping(value = "AddSeal/Word3", method = RequestMethod.GET)
    public ModelAndView showWord3(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);

        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal/test3.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal/Word3");
        return mv;
    }

    @RequestMapping(value = "AddSeal/Word4", method = RequestMethod.GET)
    public ModelAndView showWord4(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        //添加自定义按钮
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("添加印章位置", "InsertSealPos()", 2);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal/test4.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal/Word4");
        return mv;
    }

    @RequestMapping(value = "AddSeal/Word5", method = RequestMethod.GET)
    public ModelAndView showWord5(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("盖章到印章位置", "AddSealByPos()", 2);

        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal/test5.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal/Word5");
        return mv;
    }

    @RequestMapping(value = "AddSeal/Word6", method = RequestMethod.GET)
    public ModelAndView showWord6(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("盖章到印章位置", "AddSealByPos()", 2);

        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal/test6.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal/Word6");
        return mv;
    }

    @RequestMapping(value = "AddSeal/Word7", method = RequestMethod.GET)
    public ModelAndView showWord7(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("盖章到印章位置", "AddSealByPos()", 2);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal/test7.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal/Word7");
        return mv;
    }

    @RequestMapping(value = "AddSeal/Word8", method = RequestMethod.GET)
    public ModelAndView showWord8(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal/test8.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal/Word8");
        return mv;
    }

    @RequestMapping(value = "AddSeal/Word9", method = RequestMethod.GET)
    public ModelAndView showWord9(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);

        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal/test9.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal/Word9");
        return mv;
    }

    @RequestMapping(value = "AddSeal/Word10", method = RequestMethod.GET)
    public ModelAndView showWord10(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除指定印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("清除所有印章", "DeleteAllSeal()", 21);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSeal/test10.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSeal/Word10");
        return mv;
    }

    @RequestMapping(value = "AddSign/Word1", method = RequestMethod.GET)
    public ModelAndView showWord11(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

//添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "InsertHandSign()", 3);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSign/test1.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSign/Word1");
        return mv;
    }


    @RequestMapping(value = "AddSign/Word2", method = RequestMethod.GET)
    public ModelAndView showWord12(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "InsertHandSign()", 2);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSign/test2.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSign/Word2");
        return mv;
    }

    @RequestMapping(value = "AddSign/Word3", method = RequestMethod.GET)
    public ModelAndView showWord13(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "InsertHandSign()", 3);

        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSign/test3.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSign/Word3");
        return mv;
    }

    @RequestMapping(value = "AddSign/Word4", method = RequestMethod.GET)
    public ModelAndView showWord14(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "AddHandSign()", 3);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSign/test4.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSign/Word4");
        return mv;
    }

    @RequestMapping(value = "AddSign/Word5", method = RequestMethod.GET)
    public ModelAndView showWord15(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
//添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "InsertHandSign()", 3);

        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertSeal/Word/AddSign/test5.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertSeal/Word/AddSign/Word5");
        return mv;
    }


    @RequestMapping("AddSeal/save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "InsertSeal/Word/AddSeal/" + fs.getFileName());
        fs.close();
    }

    @RequestMapping("AddSign/save")
    public void save2(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "InsertSeal/Word/AddSign/" + fs.getFileName());
        fs.close();
    }

}
