package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.Map;

@RestController
@RequestMapping(value = "/FileMakerSingle/")
public class FileMakerSingleController {
    private String dir = ResourceUtils.getURL("classpath:").getPath() + "static/doc/";

    public FileMakerSingleController() throws FileNotFoundException {
    }

    @RequestMapping(value = "index", method = RequestMethod.GET)
    public ModelAndView showindex(HttpServletRequest request, Map<String, Object> map) {
        ModelAndView mv = new ModelAndView("FileMakerSingle/index");
        return mv;
    }

    @RequestMapping(value = "FileMaker", method = RequestMethod.GET)
    public ModelAndView showFileMaker(HttpServletRequest request, Map<String, Object> map) {
        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        WordDocument doc = new WordDocument();
        //禁用右击事件
        doc.setDisableWindowRightClick(true);
        //给数据区域赋值，即把数据填充到模板中相应的位置
        doc.openDataRegion("PO_company").setValue("北京卓正志远软件有限公司  ");
        fmCtrl.setSaveFilePage("save");
        fmCtrl.setWriter(doc);
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        fmCtrl.setFileTitle("newfilename.doc");
        fmCtrl.fillDocument("/doc/FileMakerSingle/test.doc", DocumentOpenType.Word);

        map.put("pageoffice", fmCtrl.getHtmlCode("FileMakerCtrl1"));
        map.put("type", request.getParameter("type"));

        ModelAndView mv = new ModelAndView("FileMakerSingle/FileMaker");
        return mv;
    }


    @RequestMapping(value = "OpenWord", method = RequestMethod.GET)
    public ModelAndView OpenWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        //设置服务器页面
        poCtrl.setCustomToolbar(false);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        poCtrl.webOpen("/doc/FileMakerSingle/maker.doc", OpenModeType.docReadOnly, "张佚名");

        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("type", request.getParameter("type"));

        ModelAndView mv = new ModelAndView("FileMakerSingle/OpenWord");
        return mv;
    }


    @RequestMapping("DownWord")
    public void DownWord(HttpServletRequest request, HttpServletResponse response) throws IOException {
        String filePath =dir + "FileMakerSingle/maker.doc";
        int fileSize =(int)new File(filePath).length();

        response.reset();
        response.setContentType("application/msword"); // application/x-excel, application/ms-powerpoint, application/pdf
        response.setHeader("Content-Disposition", "attachment; filename=maker.doc");
        response.setContentLength(fileSize);

        OutputStream outputStream = response.getOutputStream();
        InputStream inputStream = new FileInputStream(filePath);
        byte[] buffer = new byte[10240];
        int i = -1;
        while ((i = inputStream.read(buffer)) != -1) {
            outputStream.write(buffer, 0, i);
        }

        outputStream.flush();
        outputStream.close();
        inputStream.close();
    }




    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "FileMakerSingle/maker" + fs.getFileExtName());
        fs.close();
    }

}
