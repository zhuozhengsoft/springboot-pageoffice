package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.*;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.awt.*;
import java.util.Map;

@RestController
@RequestMapping(value = "/WordParagraph/")
public class WordParagraphController {

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        WordDocument doc = new WordDocument();
        //设置内容标题
        //创建DataRegion对象，PO_title为自动添加的书签名称,书签名称需以“PO_”为前缀，切书签名称不能重复
        //三个参数分别为要新插入书签的名称、新书签的插入位置、相关联的书签名称（“[home]”代表Word文档的第一个位置）
        DataRegion title = doc.createDataRegion("PO_title",
                DataRegionInsertType.After, "[home]");
        //给DataRegion对象赋值
        title.setValue("C#中Socket多线程编程实例\n");
        //设置字体：粗细、大小、字体名称、是否是斜体
        title.getFont().setBold(true);
        title.getFont().setSize(20);
        title.getFont().setName("黑体");
        title.getFont().setItalic(false);
        //定义段落对象
        ParagraphFormat titlePara = title.getParagraphFormat();
        //设置段落对齐方式
        titlePara.setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
        //设置段落行间距
        titlePara.setLineSpacingRule(WdLineSpacing.wdLineSpaceMultiple);
        //设置内容
        //第一段
        //创建DataRegion对象，PO_body为自动添加的书签名称
        DataRegion body = doc.createDataRegion("PO_body",
                DataRegionInsertType.After, "PO_title");
        //设置字体：粗细、是否是斜体、大小、字体名称、字体颜色
        body.getFont().setBold(false);
        body.getFont().setItalic(true);
        body.getFont().setSize(10);
        //设置中文字体名称
        body.getFont().setName("楷体");
        //设置英文字体名称
        body.getFont().setName("Times New Roman");
        body.getFont().setColor(Color.RED);
        //给DataRegion对象赋值
        body
                .setValue("是微软随着VS.net新推出的一门语言。它作为一门新兴的语言，有着C++的强健，又有着VB等的RAD特性。而且，微软推出C#主要的目的是为了对抗Sun公司的Java。大家都知道Java语言的强大功能，尤其在网络编程方面。于是，C#在网络编程方面也自然不甘落后于人。本文就向大家介绍一下C#下实现套接字（Sockets）编程的一些基本知识，以期能使大家对此有个大致了解。首先，我向大家介绍一下套接字的概念。\n");
        //创建ParagraphFormat对象
        ParagraphFormat bodyPara = body.getParagraphFormat();
        //设置段落的行间距、对齐方式、首行缩进
        bodyPara.setLineSpacingRule(WdLineSpacing.wdLineSpaceAtLeast);
        bodyPara.setAlignment(WdParagraphAlignment.wdAlignParagraphLeft);
        bodyPara.setFirstLineIndent(21);

        //第二段
        DataRegion body2 = doc.createDataRegion("PO_body2",
                DataRegionInsertType.After, "PO_body");
        body2.getFont().setBold(false);
        body2.getFont().setSize(12);
        body2.getFont().setName("黑体");
        body2
                .setValue("套接字是通信的基石，是支持TCP/IP协议的网络通信的基本操作单元。可以将套接字看作不同主机间的进程进行双向通信的端点，它构成了单个主机内及整个网络间的编程界面。套接字存在于通信域中，通信域是为了处理一般的线程通过套接字通信而引进的一种抽象概念。套接字通常和同一个域中的套接字交换数据（数据交换也可能穿越域的界限，但这时一定要执行某种解释程序）。各种进程使用这个相同的域互相之间用Internet协议簇来进行通信。\n");
        //body2.setValue("[image]../images/logo.jpg[/image]");
        ParagraphFormat bodyPara2 = body2.getParagraphFormat();
        bodyPara2.setLineSpacingRule(WdLineSpacing.wdLineSpace1pt5);
        bodyPara2.setAlignment(WdParagraphAlignment.wdAlignParagraphLeft);
        bodyPara2.setFirstLineIndent(21);

        //第三段
        DataRegion body3 = doc.createDataRegion("PO_body3",
                DataRegionInsertType.After, "PO_body2");
        body3.getFont().setBold(false);
        body3.getFont().setColor(Color.getHSBColor(0, 128, 228));
        body3.getFont().setSize(14);
        body3.getFont().setName("华文彩云");
        body3
                .setValue("套接字可以根据通信性质分类，这种性质对于用户是可见的。应用程序一般仅在同一类的套接字间进行通信。不过只要底层的通信协议允许，不同类型的套接字间也照样可以通信。套接字有两种不同的类型：流套接字和数据报套接字。\n");
        ParagraphFormat bodyPara3 = body3.getParagraphFormat();
        bodyPara3.setLineSpacingRule(WdLineSpacing.wdLineSpaceDouble);
        bodyPara3.setAlignment(WdParagraphAlignment.wdAlignParagraphLeft);
        bodyPara3.setFirstLineIndent(21);

        DataRegion body4 = doc.createDataRegion("PO_body4",
                DataRegionInsertType.After, "PO_body3");
        body4.setValue("[image]doc/logo.png[/image]");
        //body4.setValue("[word]doc/1.doc[/word]");//还可嵌入其他Word文件
        ParagraphFormat bodyPara4 = body4.getParagraphFormat();
        bodyPara4.setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);

        poCtrl.setWriter(doc);
        //设置页面保存后执行的JS函数
        poCtrl.setJsFunction_AfterDocumentSaved("SaveOK()");
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //打开Word文档
        poCtrl.webOpen("/doc/WordParagraph/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("WordParagraph/Word");
        return mv;
    }


}
