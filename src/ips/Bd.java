package ips;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.*;

public class Bd {	

	public String[][] SQL_Dev_Matriz_Env_Sql(String sql, int valores, int conecc) {
	  String matrizGastos[][] = new String[30000][40];
	  try {
	   Connection con = this.conecta(conecc);
	   java.sql.ResultSet rset = null;
	   rset = con.createStatement().executeQuery(sql);
	   int j = 1;
	   while (rset.next()) {
	    for (int g = 1; g <= valores; g++) {
	     matrizGastos[j][g] = rset.getString(g);
	    }
	    j++;
	   }
	   matrizGastos[j][1] = "FIN";
	   con.close();
	   con = null;
	   rset.close();
	   rset = null;
	   return matrizGastos;
	  } catch (Exception e) {
	   return null;
	  }
	 }

	public String SQL_Ejecuta_Accion_Update_Insert_Delete(String sql, int conecc, int id , int linea) {
	  try {
	   Connection con = this.conecta(conecc);
	   java.sql.ResultSet rset = null;
	   rset = con.createStatement().executeQuery(sql);
	   con.close();
	   con = null;
	   rset = null;
	   return "";
	  } catch (Exception e) {
	   return "Error Insert:" + e + "-" + "Id-Noticia:" + id + " Linea: " + linea;
	  }
	 }

	 public Connection conecta(int valor) {
	  try {
	   DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
	   Connection con = null;
	   if (valor == 7) {
		    //con = DriverManager.getConnection("jdbc:oracle:thin:@10.150.150.196:1521:legados", "dmips ", "123456");
		    con = DriverManager.getConnection("jdbc:oracle:thin:@Jorge-VAIO:1521:XE", "dmips", "123456");
	   }
	   //DatabaseMetaData meta = con.getMetaData();
	   //System.out.println("JDBC driver version is " + meta.getDriverVersion());
	   return con;
	  } catch (Exception e) {
	   return null;
	  }
	 }
	
}
