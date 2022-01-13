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
@RequestMapping(value = "/SaveDataAndFile/")
public class SaveDataAndFileController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        WordDocument wordDoc = new WordDocument();
        //打开数据区域，openDataRegion方法的参数代表Word文档中的书签名称
        DataRegion dataRegion1 = wordDoc.openDataRegion("PO_userName");
        //设置DataRegion的可编辑性
        dataRegion1.setEditing(true);
        DataRegion dataRegion2 = wordDoc.openDataRegion("PO_deptName");
        dataRegion2.setEditing(true);
        poCtrl.setWriter(wordDoc);
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存数据的页面
        poCtrl.setSaveDataPage("SaveData");
        //设置保存文档的页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/SaveDataAndFile/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));

        ModelAndView mv = new ModelAndView("SaveDataAndFile/Word");
        return mv;
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "SaveDataAndFile/" + fs.getFileName());
        fs.close();
    }

    @RequestMapping("SaveData")
    public void saveData(HttpServletRequest request, HttpServletResponse response) {
        com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);
        //获取提交的数值
        String dataUserName = doc.openDataRegion("PO_userName").getValue();
        String dataDeptName = doc.openDataRegion("PO_deptName").getValue();
        String companyName = doc.getFormField("txtCompany");

        /**获取到的公司名称,员工姓名,部门名称等内容可以保存到数据库,这里可以连接开发人员自己的数据库,例如连接mysql数据库。
         *Class.forName("com.mysql.jdbc.Driver");
         *Connection conn = (Connection) DriverManager.getConnection("jdbc:mysql://localhost/myDataBase", "root", "111111");
         *String sql="update user set userName='"+dataUserName+"',deptName='"+dataDeptName+"',companyName='"+companyName+"' where userId=1";
         *PreparedStatement ps = conn.prepareStatement(sql);
         *int rs = ps.executeUpdate(upsql);
         * if (rs>0) {
         *    out.println("更新成功");
         *     }
         *     else{
         *   out.println("更新失败");
         *    }
         *
         *rs.close();
         *conn.close();
         */
        doc.close();
    }


}
