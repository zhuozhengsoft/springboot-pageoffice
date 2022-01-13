package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.*;
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
@RequestMapping(value = "/FileMaker/")
public class FileMakerController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "index", method = RequestMethod.GET)
    public ModelAndView showindex(HttpServletRequest request, Map<String, Object> map) {
        map.put("url", dir + "FileMaker/");
        ModelAndView mv = new ModelAndView("FileMaker/index");
        return mv;
    }

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {

        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        String id = request.getParameter("id");

        if (id != null && id.length() > 0) {
            WordDocument doc = new WordDocument();
            //禁用右击事件
            doc.setDisableWindowRightClick(true);
            //给数据区域赋值，即把数据填充到模板中相应的位置
            doc.openDataRegion("PO_company").setValue("北京卓正志远软件有限公司  " + id);
            fmCtrl.setSaveFilePage("save?id=" + id);
            fmCtrl.setWriter(doc);
            fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
            fmCtrl.setFileTitle("newfilename.doc");
            fmCtrl.fillDocument("/doc/FileMaker/test.doc", DocumentOpenType.Word);

        }
        map.put("pageoffice", fmCtrl.getHtmlCode("FileMakerCtrl1"));
        ModelAndView mv = new ModelAndView("FileMaker/Word");
        return mv;
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        String id = request.getParameter("id");
        String err = "";
        if (id != null && id.length() > 0) {
            String fileName = "maker" + id + fs.getFileExtName();
            fs.saveToFile(dir + "FileMaker/" + fileName);
        } else {
            throw new RuntimeException("id 不能为空");
        }

        fs.close();
    }

}
