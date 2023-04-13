package com.web.automation;

import com.web.automation.web.DDTWeb;
import common.AutoLogger;
import common.ExcelReader;
import common.ExcelWriter;
import io.qameta.allure.testng.AllureTestNg;
import org.testng.annotations.*;

import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Date;

import static org.testng.Assert.assertEquals;

public class WebExcelDriver {
public DDTWeb web;
public ExcelReader excelReader;
public ExcelWriter resultExcel;
public String rootpath;
public String createdate;

@BeforeSuite
public void beforSuite(){
    Date date=new Date();
    SimpleDateFormat simpleDateFormat=new SimpleDateFormat("yyyyMMdd-HH-mm-ss");
    createdate = simpleDateFormat.format(date);
    rootpath = System.getProperty("user.dir");
//    打开所在路径下的文件
    excelReader=new ExcelReader(rootpath+"\\Cases\\qiandun.xlsx");
//    将文件复制一份，写入结果输出
    resultExcel=new ExcelWriter(rootpath+"\\Cases\\qiandun.xlsx",rootpath+"\\Cases\\Result\\result-"+createdate+"qiandun");
//    将DDTO里的ExcelWriter类实例化
    web=new DDTWeb(resultExcel);
}
//读取Excel数据驱动
@Test(dataProvider ="keyWords")
    public void webTest(String rowNo, String kong,String casename,String keywords,String param1,String param2,String param3,String param4,String param5,String param6,String k2,String k3,String k4){
//读取excel返回的内容中的行数，作为当前正在操作的行
    int No=Integer.parseInt(rowNo);
    web.nowline=No;
//    输出行数和用例名称
    System.out.println(rowNo+casename);
//通过java反射机制调用excel文件中指定的关键字方法。
       String runRes = runUIwithInvoke(keywords, param1, param2, param3, param4,param5,param6);
    System.out.println(runRes);
//    利用testNG的断言机制，判断用例是否通过
    assertEquals(runRes,"Pass");

}

//利用ExcelReader中的二维数组方法，读取Excel中的数据
@DataProvider
public Object[][] keyWords(){
    return excelReader.readAsMatrix();
}
    //调用的反射方式执行关键字的方法。
private String runUIwithInvoke(String key,String param1,String param2,String param3,String param4,String param5,String param6){
    String result="Fail";
    try {
        //基于参数名查找参数列表为空的方法。
        Method method = web.getClass().getDeclaredMethod(key);
        // invoke语法，需要输入类名以及相应的方法用到的参数
        result=method.invoke(web).toString();
        return result;
    } catch (Exception e) {
    }
    try {
//        基于一个参数的方法
        Method method = web.getClass().getDeclaredMethod(key, String.class);
        // invoke语法，需要输入类名以及相应的方法用到的参数
        result = method.invoke(web, param1).toString();
        return result;
    } catch (Exception e) {

    }
    try {
        //        基于两个参数的方法
        Method method = web.getClass().getDeclaredMethod(key, String.class, String.class);
        // invoke语法，需要输入类名以及相应的方法用到的参数
        result=method.invoke(web,param1,param2).toString();
        return result;
    } catch (Exception e) {
    }
    try {
        //        基于三个参数的方法
        Method method = web.getClass().getDeclaredMethod(key, String.class, String.class, String.class);
        // invoke语法，需要输入类名以及相应的方法用到的参数
        result=method.invoke(web,param1,param2,param3).toString();
        return result;
    } catch (Exception e) {

    }
    try {
        //        基于三个参数的方法
        Method method = web.getClass().getDeclaredMethod(key, String.class, String.class, String.class,String.class,String.class);
        // invoke语法，需要输入类名以及相应的方法用到的参数
        result=method.invoke(web,param1,param2,param3,param4,param5).toString();
        return result;
    } catch (Exception e) {

    }
    try {
        //        基于三个参数的方法
        Method method = web.getClass().getDeclaredMethod(key, String.class, String.class, String.class,String.class,String.class,String.class);
        // invoke语法，需要输入类名以及相应的方法用到的参数
        result=method.invoke(web,param1,param2,param3,param4,param5,param6).toString();
        return result;
    } catch (Exception e) {

    }
        return result;
//后面不够自己加
}

    @AfterSuite
    public void afterSuite() {
        excelReader.close();
        resultExcel.save();
    }


}
