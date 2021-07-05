package cn.nvwa.apicase.util;

import org.apache.commons.dbcp.BasicDataSource;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

public class DbcpUtil {
	
	private static BasicDataSource dataSource = null;

	/**
	 * 初始化连接池
	 * 避免初始化失败加try-catch
	 */
	static{
		try{
			//确保只有1个连接池
			if(dataSource == null){
				dataSource = new BasicDataSource();
				//连接数据库基础信息
		    	dataSource.setDriverClassName("com.mysql.cj.jdbc.Driver");
                //本地数据库
//		    	dataSource.setUrl("jdbc:mysql://localhost:3306/interface_testcases?characterEncoding=UTF-8");
//				dataSource.setUsername("root");
//				dataSource.setPassword("Hzj123456");

//                //压测环境数据库
                dataSource.setUrl("jdbc:mysql://10.100.130.37:3306/nvwa_api_testcases?characterEncoding=UTF-8");
                dataSource.setUsername("root");
                dataSource.setPassword("nvwa@passwd");
				
				//连接池配置信息
				dataSource.setInitialSize(10);
				dataSource.setMinIdle(10);
				dataSource.setMaxIdle(10);
				dataSource.setMaxActive(10);
				
				//连接池借出和客户端返回连接检查配置
				dataSource.setTestOnReturn(false);
				dataSource.setTestOnBorrow(false);
				dataSource.setMaxWait(2000);
				
				//连接池支持预编译
				dataSource.setPoolPreparedStatements(true);
			}
		}catch(Exception e){
			e.printStackTrace();
		}
	}
	
	/**
	 * @Title: getConnection   
	 * @Description: 同时刻只能有一个线程从池子借用连接
	 * @param: @return      
	 * @return: Connection      
	 * @throws
	 */
	public static synchronized Connection getConnection(){
		//初始化返回值
		Connection con = null;
		try{
			//从连接池借用连接
			con = dataSource.getConnection();
		}catch(Exception e){
			e.printStackTrace();
		}
		return con;
	}

	
	/**
	 * @Title: close   
	 * @Description: 关闭rs,ps,connection
	 * @param: @param rs
	 * @param: @param ps
	 * @param: @param connection      
	 * @return: void      
	 * @throws
	 */
	public static void close(ResultSet rs,PreparedStatement ps,Connection connection){
		try {
			if(rs != null){
				rs.close();
			}
			if(ps != null){
				ps.close();
			}
			if(connection != null){
				connection.close();
			}
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}

}
