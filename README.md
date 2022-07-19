# springboot-pageoffice
### 一、简介

​       springboot-pageoffice项目演示了在SpringBoot框架下如何使用PageOffice产品，此项目演示了PageOffice产品近90%的功能，每个PageOffice功能模块都以一个单独的Controller方式进行了展示，方便初学者学习和使用PageOffice产品。

### 二、项目环境要求

​       Intelij IDEA,jdk1.8及以上版本

### 三、项目运行准备

​      在当前服务器磁盘上新建一个PageOffice系统文件夹，例如：D:/pageoffice（此文件夹将用来存放PageOffice注册后生成的授权文件“license.lic”）。

### 四、项目运行步骤
1. 使用git clone或者直接下载项目压缩包到本地并解压缩。
2. 导入springboot-pageoffice项目到idea。
3. 打开application.properties文件，将posyspath变量的值配置成您上一步新建的PageOffice系统文件夹（例如：D:/pageoffice）
4. 如果您要测试PageOffice的电子印章功能，请拷贝当前项目根目录下的poseal.db文件到PageOffice系统文件夹下（例如：D:/pageoffice/poseal.db）
5. 运行项目：点击运行按钮即可。
6. 启动浏览器访问：http://localhost:8080/index ，即可在线打开、编辑、保存office文件 。
### 五、PageOffice序列号
​     PageOfficeV5.0标准版试用序列号：I2BFU-MQ89-M4ZZ-ZWY7K           
​     PageOfficeV5.0专业版试用序列号：DJMTF-HYK4-BDQ3-2MBUC
### 六、集成PageOffice到您的项目中的关键步骤
1. 在您项目的pom.xml中通过下面的代码引入PageOffice依赖。

   > pageoffice.jar的所有Releases版本都已上传到了Maven中央仓库，具体您要引用哪个版本请在Maven中央仓库地址中查看，建议使用Maven中央仓库中pageoffice.jar的最新版本。(Maven中央仓库中pageoffice.jar的地址：https://mvnrepository.com/artifact/com.zhuozhengsoft/pageoffice)

```xml
<dependency>
     <groupId>com.zhuozhengsoft</groupId>   
  <artifactId>pageoffice</artifactId>   
  <version>5.3.0.3</version>
</dependency>
```

2. 在您项目的启动类Application类中配置如下代码。

```java
@Bean
   public ServletRegistrationBean pageofficeRegistrationBean()  {
com.zhuozhengsoft.pageoffice.poserver.Server poserver = new com.zhuozhengsoft.pageoffice.poserver.Server();
/**如果当前项目是打成jar或者war包运行，强烈建议将license的路径更换成某个固定的绝对路径下，不要放当前项目文件夹下,为了防止每次重新发布项目导致license丢失问题。
  * 比如windows服务器下：D:/pageoffice，linux服务器下:/root/pageoffice
 */
//设置PageOffice注册成功后,license.lic文件存放的目录
poserver.setSysPath(poSysPath);//poSysPath可以在application.properties这个文件中配置，也可以直设置文件夹路径，比如：poserver.setSysPath("D:/pageoffice");
ServletRegistrationBean srb = new ServletRegistrationBean(poserver);
srb.addUrlMappings("/poserver.zz");
srb.addUrlMappings("/posetup.exe");
srb.addUrlMappings("/pageoffice.js");
srb.addUrlMappings("/jquery.min.js");
srb.addUrlMappings("/pobstyle.css");
srb.addUrlMappings("/sealsetup.exe");
return srb;
}
```

  3.新建Controller并调用PageOffice，例如：

```java
public class PageOfficeController {
@RequestMapping(value = "/Word", method = RequestMethod.GET)
  public ModelAndView showWord(HttpServletRequest request) {
  PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
  poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
  poCtrl.webOpen("/doc/test.doc", OpenModeType.docNormalEdit, "张三");
  request.setAttribute("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
  ModelAndView mv = new ModelAndView("Word.html");
   return mv;
  }
}
```

   4.新建View页面,例如：Word.html（PageOfficeCtroller返回的View页面，用来嵌入PageOffice控件)，PageOffice在View页面输出的代码如下：

```javascript
<div style="width: auto; height: 700px;" th:utext="${pageoffice}">
```

  5.在要打开文件的页面的head标签中先引用pageoffice.js文件后，再调POBrowser.openWindowModeless()方法打开文件，例如：

```
<!--pageoffice.js的引用路径来自于第2步的项目启动类中的配置路径,一般将此js配置到了当前项目的根目录下 -->
<script type="text/javascript" src="pageoffice.js"></script>

<!--openWindowModeless()方法的第一个参数指向的url路径是指调用pageoffice打开文件的controller路径,比如下面的"SimpleWord/Word"-->
<a href="javascript:POBrowser.openWindowModeless('SimpleWord/Word', 'width=1050px;height=900px;');">最简单在线打开保存Word文件(URL地址方式)</a>
```

### 七、电子印章功能说明

​     如果您的项目要用到PageOffice自带电子印章功能，请按下面的步骤进行操作。

​     1.请在您的web项目的启动类Application类中加上印章功能相关配置代码，具体代码请参考当前项目springboot-pageoffice启动类中电子印章功能配置代码，此处不再赘述。

> ​    注意：adminSeal.setSysPath(poSysPath)中的poSysPath指向的地址必须是您当前项目license.lic文件所在的目录。

​    2.请拷贝当前项目根目录下poseal.db文件到您的web项目的license.lic文件所在的文件夹下。

​    3.请参考当前项目的[一、15、演示加盖印章和签字功能（以Word为例）](http://localhost:8080/InsertSeal/index)  示例代码进行盖章功能调用。

### 八、 PageOffice开发帮助

​     1.[Java API文档](https://www.zhuozhengsoft.com/help/java3/index.html) 

​     2.[JS API文档](https://www.zhuozhengsoft.com/help/js3/index.html)  

​     3.[PageOffice从入门到精通](https://www.kancloud.cn/pageoffice_course_group/pageoffice_course/646953)

​     技术支持：https://www.zhuozhengsoft.com/Technical/

### 九、联系我们

​   卓正官网：[https://www.zhuozhengsoft.com](http://www.zhuozhengsoft.com)

​   联系电话：400-6600-770  

   QQ: 800038353