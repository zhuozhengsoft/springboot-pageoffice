package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;

import java.util.Map;

@RestController
@RequestMapping(value = "/ExcelDisableRight/")
public class ExcelDisableRightController {

    @RequestMapping(value = "Excel", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        Workbook workBoook = new Workbook();
        workBoook.setDisableSheetRightClick(true);//禁止当前工作表鼠标右键
        //workBoook.setDisableSheetDoubleClick(true);//禁止当前工作表鼠标双击
        //workBoook.setDisableSheetSelection(true);//禁止在当前工作表中选择内容
        poCtrl.setWriter(workBoook);

        //打开Word文档
        poCtrl.webOpen("/doc/ExcelDisableRight/test.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("ExcelDisableRight/Word");
        return mv;
    }


}
