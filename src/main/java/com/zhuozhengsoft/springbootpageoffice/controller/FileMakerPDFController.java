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
@RequestMapping(value = "/FileMakerPDF/")
public class FileMakerPDFController {
    private String dir = ResourceUtils.getURL("classpath:").getPath() + "static/doc/";

    public FileMakerPDFController() throws FileNotFoundException {
    }

    @RequestMapping(value = "index", method = RequestMethod.GET)
    public ModelAndView showindex(HttpServletRequest request, Map<String, Object> map) {
        ModelAndView mv = new ModelAndView("FileMakerPDF/index");
        return mv;
    }

    @RequestMapping(value = "FileMaker", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
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

        fmCtrl.fillDocumentAsPDF("/doc/FileMakerPDF/template.doc", DocumentOpenType.Word, "a.pdf");
        map.put("pageoffice", fmCtrl.getHtmlCode("FileMakerCtrl1"));
        map.put("type", request.getParameter("type"));
        ModelAndView mv = new ModelAndView("FileMakerPDF/FileMaker");
        System.out.println();
        return mv;
    }
    @RequestMapping(value = "OpenPDF", method = RequestMethod.GET)
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
        pdfCtrl.webOpen("/doc/FileMakerPDF/template.pdf");

        map.put("pageoffice", pdfCtrl.getHtmlCode("PDFCtrl1"));
        ModelAndView mv = new ModelAndView("FileMakerPDF/pdf");
        return mv;
    }

    

    @RequestMapping("DownPDF")
    public void DownPDF(HttpServletRequest request, HttpServletResponse response) throws IOException {
        String filePath = dir + "FileMakerPDF/template.pdf";
        int fileSize = (int) new File(filePath).length();

        response.reset();
        response.setContentType("application/pdf"); // application/x-excel, application/ms-powerpoint, application/pdf
        response.setHeader("Content-Disposition", "attachment; filename=template.pdf");
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
        fs.saveToFile(dir + "FileMakerPDF/" + fs.getFileName());
        fs.close();
    }

}
