package com.web.automation.web;

import com.google.common.io.Files;
import common.AutoLogger;
import common.ExcelWriter;
import org.openqa.selenium.*;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import web.RobotUtils;
import webDriver.FFDriver;
import webDriver.GoogleDriver;
import webDriver.IEDriver;

import javax.print.DocFlavor;
import javax.xml.bind.Element;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.TimeUnit;

public class DDTWeb {
    private WebDriver driver;
//    创建ExcelWriter对象
    public ExcelWriter resultcase;
//    写入的列
    public final int ResultCol=12;
//    写入操作的时候，需要告诉wrtier写入的行号
    public int nowline=0;
    public int winTime;
    //构造方法中，传递resultwriter，完成excelwriter的赋值操作。其它方法中调用完成写入。
    public void setLine(int rowNo){
        nowline=rowNo;
    }
//    存在一个有参的构造方法，写一个无参的
    public DDTWeb(){
    }
    public DDTWeb(ExcelWriter excelWriter){
        resultcase=excelWriter;
    }
    public WebDriver getWebDriver() {
        return driver;
    }

    public void setWebDriver(WebDriver Driver) {
        this.driver = driver;
    }
    /*
    * type指的是使用的浏览器
    * webDriverPath指的是浏览器驱动路径（已经创建了一个包webDriverExe，将浏览器驱动放入即可）
    * seconds:为隐式等待最长时间
    * */
    public String openBrowser(String type,String webDriverPath,String seconds){
        try {
            switch (type){
                case "chrome":
                    GoogleDriver googleDriver = new GoogleDriver(webDriverPath);
                    driver=googleDriver.getdriver();
                    //窗口最大化
                    driver.manage().window().maximize();
                    winTime=Integer.parseInt(seconds);
                    driver.manage().timeouts().implicitlyWait(winTime, TimeUnit.SECONDS);
                    AutoLogger.log.info("****************浏览器启动************");
                    break;
                case "firefox":
                    FFDriver ffDriver=new FFDriver("",webDriverPath);
                    driver=ffDriver.getdriver();
                    //窗口最大化
                    driver.manage().window().maximize();
                    winTime=Integer.parseInt(seconds);
                    driver.manage().timeouts().implicitlyWait(winTime,TimeUnit.SECONDS);
                    AutoLogger.log.info("****************浏览器启动************");
                    break;
                case "ie":
                    IEDriver ieDriver=new IEDriver(webDriverPath);
                    driver=ieDriver.driver;
                    //窗口最大化
                    driver.manage().window().maximize();
                    winTime=Integer.parseInt(seconds);
                    driver.manage().timeouts().implicitlyWait(winTime,TimeUnit.SECONDS);
                    AutoLogger.log.info("****************浏览器启动************");
                    break;
                case "edge":
                    System.setProperty("webdriver.edge.driver",webDriverPath);
                    driver = new EdgeDriver();
                    //窗口最大化
                    driver.manage().window().maximize();
                    winTime=Integer.parseInt(seconds);
                    driver.manage().timeouts().implicitlyWait(winTime, TimeUnit.SECONDS);
                    AutoLogger.log.info("****************浏览器启动************");
                    break;
                default:
                    GoogleDriver ggDriver = new GoogleDriver(webDriverPath);
                    driver=ggDriver.getdriver();
                    //窗口最大化
                    driver.manage().window().maximize();
                    winTime=Integer.parseInt(seconds);
                    driver.manage().timeouts().implicitlyWait(winTime, TimeUnit.SECONDS);
                    AutoLogger.log.info("****************浏览器启动************");
                    break;
            }
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            e.printStackTrace();
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }


    }
    //打开网址
    public String getUrl(String url){
        try {
            driver.get(url);
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            e.printStackTrace();
            AutoLogger.log.error("打开浏览器失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }
    //强制等待
    public String halt(String seconds){
        try {
            double time = Double.parseDouble(seconds);
            Thread.sleep((long)(time*1000));
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (InterruptedException e) {
            e.printStackTrace();
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }
    //显示等待，运用元素定位的方法，找到相应的文本值value
    /*
    * xpath:元素路径
    * value：元素值
    * */
public String waitTextFind(String xpath,String value){
    try {
       WebDriverWait webDriverWait=new WebDriverWait(driver,10);
       webDriverWait.until(ExpectedConditions.textToBe(By.xpath(xpath),value));
        resultcase.writeCell(nowline,ResultCol,"Pass");
        return "Pass";
    } catch (Exception e) {
        e.printStackTrace();
        AutoLogger.log.error("未找到"+xpath+"下的"+value);
        resultcase.writeCell(nowline,ResultCol,"Fail");
        return "Fail";
    }
}
/*
* 利用xpath来定位点击的方法
* */
public String clickByXpath(String xpath){
    try {
        driver.findElement(By.xpath(xpath)).click();
        resultcase.writeCell(nowline,ResultCol,"Pass");
        return "Pass";
    } catch (Exception e) {
        AutoLogger.log.error("点击"+xpath+"失败");
        resultcase.writeCell(nowline,ResultCol,"Fail");
        return "Fail";
    }
}
    /*
     * 利用xpath来输入的方法
     *
     */
    public String setKeyByXpath(String xpath,String value){
        try {
            WebElement input = driver.findElement(By.xpath(xpath));
            input.clear();
            input.sendKeys(value);
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            e.printStackTrace();
            AutoLogger.log.error("定位"+xpath+"进行输入失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }

    }
//    判断是否需要登录，如果需要，则进行登录
    public String CkickLogin(String useName ,String passWord,String agreement,String loginBytton,String yourUseName,String yourPassWord){
        boolean needLogin=false;
        WebElement findUseName=null;
        WebElement findPassWord=null;
        WebElement findLoginBytton=null;
        WebElement findAgreement=null;
        try {
//            先找元素
            findUseName = driver.findElement(By.xpath(useName));
            findPassWord = driver.findElement(By.xpath(passWord));
            findAgreement=driver.findElement(By.xpath(agreement));
            findLoginBytton = driver.findElement(By.xpath(loginBytton));
            needLogin = true;
        }catch (Exception e){

        }
        try {
//            如果找到了，走这个
            if (needLogin){
                findUseName.clear();
                findUseName.sendKeys(yourUseName);
                findPassWord.clear();
                findPassWord.sendKeys(yourPassWord);
                findAgreement.click();
                findLoginBytton.click();
                resultcase.writeCell(nowline,ResultCol,"Pass");
                return "Pass";
            }else {
//                找不到，走这里，不用进行登录，但是功能用例还是通过
                AutoLogger.log.info("用户已经登录过了，不需要登录");
                resultcase.writeCell(nowline,ResultCol,"Pass");
                return "Pass";
            }

        }catch (Exception e){
//            出错后抛错
            e.printStackTrace();
            AutoLogger.log.error("输入密码定位失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";

        }


    }
    /*
    * 如果不用xpath，也可以用其它方法来定位进行点击操作
    * */
    public String click(String type,String method){
        try {
            switch (type){
                case "id":
                    driver.findElement(By.id(method)).click();
                    break;
                case "name":
                    driver.findElement(By.name(method)).click();
                    break;
                case "tag":
                    driver.findElement(By.tagName(method)).click();
                    break;
                case "classname":
                    driver.findElement(By.className(method)).click();
                    break;
                case "xpath":
                    driver.findElement(By.xpath(method)).click();
                    break;
                case "css":
                    driver.findElement(By.xpath(method)).click();
                    break;
                default:
                    driver.findElement(By.xpath(method)).click();
                    break;
            }
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            e.printStackTrace();
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }
    /*
     * 如果不用xpath，也可以用其它方法来定位进行输入操作
     * */
    public String setKey(String type,String method,String value){
        try {
            switch (type){
                case "id":
                    driver.findElement(By.id(method)).sendKeys(value);
                    break;
                case "name":
                    driver.findElement(By.name(method)).sendKeys(value);
                    break;
                case "tag":
                    driver.findElement(By.tagName(method)).sendKeys(value);
                    break;
                case "classname":
                    driver.findElement(By.className(method)).sendKeys(value);
                    break;
                case "xpath":
                    driver.findElement(By.xpath(method)).sendKeys(value);
                    break;
                case "css":
                    driver.findElement(By.xpath(method)).sendKeys(value);
                    break;
                default:
                    driver.findElement(By.xpath(method)).sendKeys(value);
                    break;
            }
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            e.printStackTrace();
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }
    //浏览器回退操作
    public String rollBack(){
        try {
            driver.navigate().back();
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            AutoLogger.log.error("浏览器退回失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }
//    浏览器前进操作
    public String advance(){
        try {
            driver.navigate().forward();
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            AutoLogger.log.error("浏览器前进失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }
//    鼠标悬停操作，需要用到工具类
    public String hover(String xpath){
        try {
           Actions actions=new Actions(driver);
           actions.moveToElement(driver.findElement(By.xpath(xpath))).perform();
            AutoLogger.log.error("悬停到指定元素");
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            e.printStackTrace();
            AutoLogger.log.error(xpath+"悬停失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }
//    使用AutoIt插件进行文件上传,AuroIt文件自己生成
    public String uploadFiles(String filePath){
        try {
            Runtime.getRuntime().exec(filePath);
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            e.fillInStackTrace();
            AutoLogger.log.error("上传文件失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }
//    删除所有cookie
    public String deleteCookies(){
        try {
            driver.manage().deleteAllCookies();
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            e.fillInStackTrace();
            AutoLogger.log.error("删除失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }
//    删除指定cookie
    public String deleteOneCookie(Cookie cookie){
        try {
            driver.manage().deleteCookie(cookie);
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            e.fillInStackTrace();
            AutoLogger.log.error("删除失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }


    /*
    下拉框处理，运用到select类
    利用输入的method（text或value）来使用不同的选择方法
    */
    public String select(String xpath,String method,String value){
        try{
        Select select=new Select(driver.findElement(By.xpath(xpath)));
        switch (method){
            case "text":
                select.selectByVisibleText(value);
                break;
            case"value":
                select.selectByValue(value);
                    break;
                }
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e ) {
                    e.printStackTrace();
                    AutoLogger.log.error("下拉框选择失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
                }
        }
    public String selectByValue(String xpath,String value){
        try{
            Select select=new Select(driver.findElement(By.xpath(xpath)));
            select.selectByValue(value);
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e ) {
            AutoLogger.log.error("下拉框选择失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }
//     利用title实现窗口切换
    public String switchWindowByTitle(String title){
                try {
                    //获取全部页面的windowhandles
                    Set<String> windowHandles = driver.getWindowHandles();
//            进行循环
                    for (String windowHandle : windowHandles) {
//             循环一次，切换一次窗口
                        driver.switchTo().window(windowHandle);
//                如果切换过来的窗口与预期的title值一样，就停止
                        if(driver.getTitle().equals(title)){
                            break;
                        }
                    }
                    resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            e.printStackTrace();
            AutoLogger.log.error("窗口切换失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }


    }
//关闭第一个老窗口
    public String closeOldWin(){
        List<String>handleList=new ArrayList<>();
//        返回一个句柄集合
        Set<String>handler=driver.getWindowHandles();
//        循环获取集合里面的句柄，保存到List数组handles里面
        Iterator iterator=handler.iterator();
//        使用迭代器进行循环
        while (iterator.hasNext()){
            handleList.add(iterator.next().toString());
        }
        //关闭第一个窗口
        driver.close();
        try {
            driver.switchTo().window(handleList.get(1));
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            AutoLogger.log.error("关闭旧窗口切换到新窗口失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }

    }
//    关闭第二个新的窗口
    public String closeNewWin(){
        List<String>handleList=new ArrayList<>();
//        返回一个句柄集合
        Set<String>handler=driver.getWindowHandles();
//        循环获取集合里面的句柄，保存到List数组handles里面
        Iterator iterator=handler.iterator();
//        使用迭代器进行循环
        while (iterator.hasNext()){
            handleList.add(iterator.next().toString());
        }
        //关闭第一个窗口
        try {
            driver.switchTo().window(handleList.get(1));
//            关闭第二个窗口
            driver.close();
            driver.switchTo().window(handleList.get(0));
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            AutoLogger.log.error("关闭旧窗口切换到新窗口失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }

    }

    //切换到工作的Iframe页面
    public String switchToIframe(String xpath){
        try {
            driver.switchTo().frame(driver.findElement(By.xpath(xpath)));
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            e.printStackTrace();
            AutoLogger.log.error("切换到iframe失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }

    }
    //切换到父级iframe页面
    public String switchToParentIframe(){
        try {
            driver.switchTo().parentFrame();
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            e.printStackTrace();
            AutoLogger.log.error("切换到父级Iframe页面失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }
//    切换到根iframe页面
    public String switchToDefaultIframe(){
        try {
            driver.switchTo().defaultContent();
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            e.printStackTrace();
            AutoLogger.log.error("切换到根Iframe页面失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }
    //利用js进行滚动操作
    public String scrollWindowStraight(String height){
        JavascriptExecutor js=(JavascriptExecutor) driver;
        try {
            String jsCmd = "window.scrollTo(0," + height + ")";
            js.executeScript(jsCmd);
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            AutoLogger.log.error("滚动条滚动失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }

    }
/**
 * 将元素定位表达式通过xpath进行定位，然后再调用要对元素进行的操作，注意是js方法的操作。
 * @param method  调用的方法 注意根据情况带上小括号进行调用  比如click() 或者 innerText="xxx"
 * @param xpath
 */
public String runJsWithElement(String method,String xpath){
    try {
        WebElement element=driver.findElement(By.xpath(xpath));
        JavascriptExecutor runner=(JavascriptExecutor)driver;
        runner.executeScript("arguments[0]."+method,element);
        resultcase.writeCell(nowline,ResultCol,"Pass");
        return "Pass";
    } catch (Exception e) {
        AutoLogger.log.error("运行JS脚本失败");
        resultcase.writeCell(nowline,ResultCol,"Fail");
        return "Fail";
    }
}
//获取页面标题
    public String getTitle(){
        try {
            driver.getTitle();
            resultcase.writeCell(nowline,ResultCol,"Pass");
            return "Pass";
        } catch (Exception e) {
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }
    /**
     * 断言标题中包含指定内容
     *
     * @param target 标题包含的内容
     */
    public String assertTitleContains(String target) {
        String result = getTitle();
        if (result.contains(target)) {
            AutoLogger.log.info("测试成功！");
            resultcase.writeCell(nowline, ResultCol, "Pass");
            return "Pass";
        } else {
            AutoLogger.log.info("测试失败！");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }

    /*
* 获取页面元素信息是否与预期值预期值来做断言
* */
    public String assertGetText(String xpath,String value){
        String text = driver.findElement(By.xpath(xpath)).getText();
        boolean result=false;
        if (text.equals(value)){
            System.out.println("测试成功");
            resultcase.writeCell(nowline, ResultCol, "Pass");
            return "Pass";
        }else {
            System.out.println("测试失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }

    }
    public String assertPageText(String value){
        String pageSource = driver.getPageSource();
        if (pageSource.contains(value)){
            System.out.println("测试成功");
            resultcase.writeCell(nowline, ResultCol, "Pass");
            return "Pass";
        }else {
            System.out.println("测试失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }
//    将文件截图后，保存至logs/screenShort下面
    public void takeScreen(String method){
        TakesScreenshot screenshot=(TakesScreenshot)driver;
        File screenshotAs = screenshot.getScreenshotAs(OutputType.FILE);
        File filePath = new File("logs/screenShot/" + method + createTime("MMdd-HH-ss") + ".png");
        try {
            Files.copy(screenshotAs,filePath);
        } catch (Exception e) {
            e.printStackTrace();
            AutoLogger.log.error("截图失败，请检查一下代码情况");
        }

    }
    /**
     * 方法用于生成指定格式的日期字符串。
     * @param format  指定格式   yyyy表示年，MM 月 dd天  HH 小时  mm 分钟  ss秒 sss 毫秒。
     */
    public String createTime(String format){
        Date date=new Date();
        SimpleDateFormat dateFormat=new SimpleDateFormat(format);
        String result = dateFormat.format(date);
        return result;

    }
//    关闭浏览器，释放资源
    public String closeBrowser(){
        try {
            driver.quit();
            resultcase.writeCell(nowline, ResultCol, "Pass");
            return "Pass";
        } catch (Exception e) {
            AutoLogger.log.error("浏览器关闭失败");
            resultcase.writeCell(nowline,ResultCol,"Fail");
            return "Fail";
        }
    }

    /*****************************************************************用鼠标模拟用户操作******************************************************/
//很少用
    public String robot(int x,int y){
    try {
        RobotUtils robot=new RobotUtils();
        robot.moveToclick(x,y);
        resultcase.writeCell(nowline, ResultCol, "Pass");
        return "Pass";
    } catch (Exception e) {
        AutoLogger.log.error("鼠标点击"+x+y+"失败");
        resultcase.writeCell(nowline, ResultCol, "Fail");
        return "Fail";
    }
}
}
