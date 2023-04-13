package webDriver;

import java.io.File;
import java.io.IOException;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeDriverService;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;

public class GoogleDriver { // Chrome浏览器驱动类
	private WebDriver driver = null;

	public GoogleDriver(String driverpath) {
		// 设置 chrome 的路径
		System.setProperty("webdriver.chrome.driver", driverpath);
		//设置临时环境变量，指定chrome使用静默模式，减少日志输出量
		System.setProperty("webdriver.chrome.silentOutput", "true");
		/**
		 * Chromeoption对象可以为chrome启动时定制许多参数，需要用到更多参数可以查阅chromeoptions相关说明。
		 */
		// chrome浏览器参数对象
		ChromeOptions option = new ChromeOptions();
		//设置自动化启动时，不显示正在受到自动化软件控制的提示栏
		option.setExperimentalOption("excludeSwitches", new String[] {"enable-automation","load-extension"});
		/**
		 * 加载chrome用户文件，这里使用的是浏览器默认存储的用户文件目录。 在chrome浏览器里通过地址栏里输入chrome://version
		 * 进行访问，能够看到个人资料路径 注意个人资料路径中复制时，只需要到User Data这一级，不需要Default这一级
		 * 使用时会和手动打开的浏览器冲突，注意不要同时打开。
		 */
		option.addArguments("--user-data-dir=C:\\Users\\Administrator\\AppData\\Local\\Google\\Chrome\\UserData");
		// 也可以将浏览器路径下的User Data目录复制一份放到其它位置进行引用，这样不会和手动打开的浏览器产生冲突。
		//option.addArguments("--user-data-dir=D:\\chromeData\\copyData");
		// 最大化浏览器窗口
		option.addArguments("--start-maximized");

		try { // 创建一个 Chrome 的浏览器实例
			this.driver = new ChromeDriver(option);
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("log--error：创建driver失败！！");
		}

	}

	public WebDriver getdriver() {
		return this.driver;
	}
}