package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import com.zhuozhengsoft.springbootpageoffice.util.GetDirPathUtil;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@RestController
@RequestMapping(value = "/SetDrByUserWord/")
public class SetDrByUserWordController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "index", method = RequestMethod.GET)
    public ModelAndView showindex(HttpServletRequest request, Map<String, Object> map) {
        ModelAndView mv = new ModelAndView("SetDrByUserWord/index");
        return mv;
    }

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        String userName = request.getParameter("userName");
        //***************************卓正PageOffice组件的使用********************************
        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion dTitle = doc.openDataRegion("PO_title");
        //给数据区域赋值
        dTitle.setValue("某公司第二季度产量报表");
        //设置数据区域可编辑性
        dTitle.setEditing(false);//数据区域不可编辑

        DataRegion dA1 = doc.openDataRegion("PO_A_pro1");
        DataRegion dA2 = doc.openDataRegion("PO_A_pro2");
        DataRegion dB1 = doc.openDataRegion("PO_B_pro1");
        DataRegion dB2 = doc.openDataRegion("PO_B_pro2");

        //根据登录用户名设置数据区域可编辑性
        //A部门经理登录后
        if (userName.equals("zhangsan")) {
            userName = "A部门经理";
            dA1.setEditing(true);
            dA2.setEditing(true);
            dB1.setEditing(false);
            dB2.setEditing(false);
        }
        //B部门经理登录后
        else {
            userName = "B部门经理";
            dB1.setEditing(true);
            dB2.setEditing(true);
            dA1.setEditing(false);
            dA2.setEditing(false);
        }

        poCtrl.setWriter(doc);
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
        poCtrl.setMenubar(false);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/SetDrByUserWord/test.doc", OpenModeType.docSubmitForm, userName);
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("userName", userName);
        ModelAndView mv = new ModelAndView("SetDrByUserWord/Word");
        return mv;
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "SetDrByUserWord/" + fs.getFileName());
        fs.close();
    }

}
