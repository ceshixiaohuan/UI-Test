package common;

import org.testng.IReporter;
import org.testng.ISuite;
import org.testng.xml.XmlSuite;

import javax.mail.*;
import javax.mail.internet.*;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Properties;

public class MyReporter implements IReporter {

    @Override
    public void generateReport(List<XmlSuite> xmlSuites, List<ISuite> suites, String outputDirectory) {
        // 在这里收集测试结果信息，例如Allure报告的JSON文件

        // 创建Session对象
        Properties properties = new Properties();
        properties.setProperty("mail.smtp.host", "smtp.gmail.com");
        properties.setProperty("mail.smtp.port", "587");
        properties.setProperty("mail.smtp.auth", "true");
        properties.setProperty("mail.smtp.starttls.enable", "true");
        Session session = Session.getInstance(properties, new Authenticator() {
            @Override
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication("username", "password");
            }
        });

        // 创建邮件对象
        try {
            MimeMessage message = new MimeMessage(session);
            message.setFrom(new InternetAddress("sender@example.com"));
            message.setRecipients(Message.RecipientType.TO, InternetAddress.parse("recipient@example.com"));
            message.setSubject("Test Results");
            message.setText("Please find attached the test results.");

            // 创建邮件附件
            Multipart multipart = new MimeMultipart();
            MimeBodyPart attachmentBodyPart = new MimeBodyPart();
            attachmentBodyPart.attachFile(new File("/path/to/allure-report.html"));
            attachmentBodyPart.setDescription("Allure Report");
            attachmentBodyPart.setFileName("allure-report.html");
            multipart.addBodyPart(attachmentBodyPart);
            message.setContent(multipart);

            // 发送邮件
            Transport.send(message);
            System.out.println("Email sent successfully.");
        } catch (MessagingException | IOException e) {
            System.out.println("Email could not be sent.");
            e.printStackTrace();
        }
    }
}
