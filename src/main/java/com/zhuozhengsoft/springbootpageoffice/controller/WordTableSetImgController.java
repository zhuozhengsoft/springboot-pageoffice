package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.Table;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@RestController
@RequestMapping(value = "/WordTableSetImg/")
public class WordTableSetImgController {

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        WordDocument doc = new WordDocument();
        Table table1 = doc.openDataRegion("PO_T001").openTable(1);

        table1.openCellRC(1, 1).setValue("[image]/doc/WordTableSetImg/wang.gif[/image]");
        int dataRowCount = 5;//需要插入数据的行数
        int oldRowCount = 3;//表格中原有的行数
        // 扩充表格
        for (int j = 0; j < dataRowCount - oldRowCount; j++) {
            table1.insertRowAfter(table1.openCellRC(2, 5));  //在第2行的最后一个单元格下插入新行
        }
        // 填充数据
        int i = 1;
        while (i <= dataRowCount) {
            table1.openCellRC(i, 2).setValue("AA" + String.valueOf(i));
            table1.openCellRC(i, 3).setValue("BB" + String.valueOf(i));
            table1.openCellRC(i, 4).setValue("CC" + String.valueOf(i));
            table1.openCellRC(i, 5).setValue("DD" + String.valueOf(i));
            i++;
        }
        poCtrl.setWriter(doc);
        //打开Word文档
        poCtrl.webOpen("/doc/WordTableSetImg/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("WordTableSetImg/Word");
        return mv;
    }

}
