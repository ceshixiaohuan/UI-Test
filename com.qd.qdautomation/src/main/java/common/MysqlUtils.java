package common;

import java.sql.*;
import java.util.*;


public class MysqlUtils {

    //连接的成员变量，后面的方法都会使用这个创建的连接
    private Connection myconnector;

    //建立连接
    public void createConnector(){
        //1、使用mysql连接驱动jdbc
        try {
            Properties royp=new Properties();
            //读取resources中的文件内容
            royp.load(MysqlUtils.class.getResourceAsStream("/shopdb.properties"));
            String jdbcurl=royp.getProperty("jdbcurl");
            String username=royp.getProperty("username");
            String classname=royp.getProperty("driverclass");
            String password=royp.getProperty("password");
            Class.forName(classname);
            //2、输入ip、端口、数据库名、用户名密码 创建连接
            //url格式： 协议jdbc:mysql:// IP : 端口/连接的数据库名 ？参数配置
            myconnector= DriverManager.getConnection(jdbcurl,username,password);
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("mysql数据库连接失败");
        }
    }

    /**
     * 通过输入用户名和密码进行查询，返回查询到的所有结果。
     * @return   返回查询到的所有结果 ，格式为List<Map<String,String>>  每个map就是一行数据。
     * @throws SQLException
     */
    public List<Map<String,String>> queryResult(String sql)  {
        try {
            //3、创建查询
            Statement roystatement = myconnector.createStatement();
            //5、执行查询
            ResultSet resultSet = roystatement.executeQuery(sql);
            ResultSetMetaData thead = resultSet.getMetaData();
            //使用一个list，把所有的map都存起来。
            List<Map<String,String>> resultlist=new ArrayList<>();
            //使用循环，读取resultSet中的每一行，通过next()取值。
            while(resultSet.next()){
                //循环中，每获取一次数据，创建一个新的map来存这一行的值
                //一行信息存储到map中
                Map<String,String> resultmap=new HashMap<>();
                //遍历表头信息的每一列，获取列名和值。
                for(int column=1;column<=thead.getColumnCount();column++){
                    resultmap.put(thead.getColumnName(column),resultSet.getString(column));
                }
                //将每一行的数据map存到list里面
                resultlist.add(resultmap);
            }
            return resultlist;
        } catch (SQLException throwables) {
            throwables.printStackTrace();
            return null;
        }
    }

}
