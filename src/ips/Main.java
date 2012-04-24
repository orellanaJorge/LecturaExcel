package ips;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import ips.Bd;
import ips.DevuelveDetalleTAF_Detall;

public class Main{

	
public static void main(String[] args){	
	try{
		
		System.out.println("Inicio");
		//ProcesoDetalle();		
		
		
		ProcesoConductor("fechafak");
		
		
		//	hojaConductor - Updatea 1 si tiene conductor y retorna su hoja , tambien updatea Hoja Detalle, hoja anexo			
		//ProcesoConductor("hojaConductor");
		
		//ProcesoConductor("ord");
		
		//ProcesoConductor("folio");
		
		//ProcesoConductor("fecha");
		
		//ProcesoConductor("totales");
		
		//ProcesoConductor("rutnomb");
		
		System.out.println("Fin");
					
	}catch(Exception e){}	
	
}

public static void ProcesoConductor(String caso){
	int h = 0;
	String insert = "";
	String fileName = "";
	DevuelveDetalleTAF_Detall back = new DevuelveDetalleTAF_Detall();
	DevuelveDetalleTAF_Conduc front = new DevuelveDetalleTAF_Conduc();
	
	String MatrizCarlos[][] = new String[15000][500];	
	String Matriz[][] = new String[15000][500];
	
	
	int valor = 0;String sql = "";
	
	if(caso.trim().equals("hojaConductor")){			
		sql = " select taf_path "; 
		sql+= " from taf_reporte ";
		sql+= " where analisar = 'S' ";
		sql+= " order by 1 ";
		
		
		
		MatrizCarlos = back.CasosParaAnalizarConduct(sql);
		
		sql = " select count(taf_path) "; 
		sql+= " from taf_reporte ";
		sql+= " where analisar = 'S' ";
		sql+= " order by 1 ";	
		
		valor = back.cuentaCasos(sql);
	}
	
	if(caso.trim().equals("ord") 
			|| caso.trim().equals("folio")
			|| caso.trim().equals("fecha")
			|| caso.trim().equals("totales")
			|| caso.trim().equals("rutnomb")
			|| caso.trim().equals("fechafak")
			){			
		sql = " select taf_path , hoja "; 
		sql+= " from taf_reporte ";
		sql+= " where conductor = 1  and nombre is null ";
		sql+= " order by 1 ";		
		MatrizCarlos = back.CasosParaAnalizar(sql);
		
		sql = " select count(taf_path) "; 
		sql+= " from taf_reporte ";
		sql+= " where  conductor = 1  and nombre is null";
		sql+= " order by 1 ";	
		valor = back.cuentaCasos(sql);
	}
	
	
	try{
	int ok = 0,x = 0;
	String cumple = "",detalle = "",error = "" ;	
	
	for(x=1; x<=valor; x++){		
		h = 0;
		Bd operaBD = new Bd();		
		fileName ="C:\\Opt\\Xls\\NuevoMaterial\\" + MatrizCarlos[x][1].trim();	
		//fileName = "C:\\Opt\\Xls\\NuevoMaterial\\125969429.xls"; 
		// hoja Conductor y estado 1 o 0 conductor, hoa anexo 
		if(caso.trim().equals("hojaConductor")){			
			Matriz = front.DevuelveConductorHoja(fileName);
		}
		
		int ghf = 0;
		String pasa = "";
		if(caso.trim().equals("ord") ){
			ghf = Integer.parseInt(MatrizCarlos[x][2].trim());
			
			if (ghf==0){
				pasa="NO";
			}else{
				pasa="SI";
				Matriz = front.DevuelveConductorOrd(fileName,ghf );
				if(Matriz[1][3] == null ){Matriz[1][3] = "";}
			    if(Matriz[1][4] == null ){Matriz[1][4] = "";}
			    if(Matriz[1][5] == null ){Matriz[1][5] = "";}
			    if(Matriz[1][6] == null ){Matriz[1][6] = "";}
			    if(Matriz[1][9] == null ){Matriz[1][9] = "0";}		      
			}
		}
		
		if(caso.trim().equals("fechafak") ){
			ghf = Integer.parseInt(MatrizCarlos[x][2].trim());
			//ghf = 1;
			if (ghf==0){
				pasa="NO";
			}else{
				pasa="SI";
				Matriz = front.DevuelveConductorFechaFak(fileName,ghf );
				if(Matriz[1][3] == null ){Matriz[1][3] = "";}
			    if(Matriz[1][4] == null ){Matriz[1][4] = "";}
			    if(Matriz[1][5] == null ){Matriz[1][5] = "";}
			    if(Matriz[1][6] == null ){Matriz[1][6] = "";}
			    if(Matriz[1][9] == null ){Matriz[1][9] = "0";}		      
			}
		}
		
			
		if(caso.trim().equals("folio") ){
			ghf = Integer.parseInt(MatrizCarlos[x][2].trim());
			if (ghf==0){
				pasa="NO";
			}else{
				pasa="SI";
				Matriz = front.DevuelveConductorFolio(fileName,ghf );
				if(Matriz[1][3] == null ){Matriz[1][3] = "";}
			    if(Matriz[1][4] == null ){Matriz[1][4] = "";}
			    if(Matriz[1][5] == null ){Matriz[1][5] = "";}
			    if(Matriz[1][6] == null ){Matriz[1][6] = "";}
			    if(Matriz[1][9] == null ){Matriz[1][9] = "0";}		      
			}
		}

		
		if(caso.trim().equals("fecha") ){
			ghf = Integer.parseInt(MatrizCarlos[x][2].trim());
			if (ghf==0){
				pasa="NO";
			}else{
				pasa="SI";
				Matriz = front.DevuelveConductorFecha(fileName,ghf );
				if(Matriz[1][3] == null ){Matriz[1][3] = "";}
			    if(Matriz[1][4] == null ){Matriz[1][4] = "";}
			    if(Matriz[1][5] == null ){Matriz[1][5] = "";}
			    if(Matriz[1][6] == null ){Matriz[1][6] = "";}
			    if(Matriz[1][9] == null ){Matriz[1][9] = "0";}		      
			}
		}
		
		if(caso.trim().equals("totales") ){
			ghf = Integer.parseInt(MatrizCarlos[x][2].trim());
			if (ghf==0){
				pasa="NO";
			}else{
				pasa="SI";
				Matriz = front.DevuelveConductorTotales(fileName,ghf );
				if(Matriz[1][3] == null ){Matriz[1][3] = "";}
			    if(Matriz[1][4] == null ){Matriz[1][4] = "";}
			    if(Matriz[1][5] == null ){Matriz[1][5] = "";}
			    if(Matriz[1][6] == null ){Matriz[1][6] = "";}
			    if(Matriz[1][9] == null ){Matriz[1][9] = "0";}		      
			}
		}
		
		if(caso.trim().equals("rutnomb") ){
			ghf = Integer.parseInt(MatrizCarlos[x][2].trim());
			if (ghf==0){
				pasa="NO";
			}else{
				pasa="SI";
				Matriz = front.DevuelveConductorRutnomb(fileName,ghf );
				if(Matriz[1][3] == null ){Matriz[1][3] = "";}
			    if(Matriz[1][4] == null ){Matriz[1][4] = "";}
			    if(Matriz[1][5] == null ){Matriz[1][5] = "";}
			    if(Matriz[1][6] == null ){Matriz[1][6] = "";}
			    if(Matriz[1][9] == null ){Matriz[1][9] = "0";}		      
			}
		}
	      
	      
	      System.out.println(x + " De " + valor);
	      
	      if (x==44){
	    	  String asdj = "";
	      }
	      
	      
	      if(caso.trim().equals("fechafak")&&pasa.trim().equals("SI")){	
    		  
    		  if(Matriz[1][1].trim().length()>0  ){
	     		 sql = "update TAF_REPORTE set fecha =  '" + Matriz[1][1].trim()+ "' where taf_path = '" +MatrizCarlos[x][1].trim() + "'"; 
	     		 insert = operaBD.SQL_Ejecuta_Accion_Update_Insert_Delete(sql, 7, 0, 0);		    	 
	     		 System.out.println("Es ok ?¡ : " + insert);
	     		 System.out.println("");
    		  }else{
    			  //System.out.println("No se Rescato FECHA");
    			  //System.out.println("");
    		  }
    		  
    		  if(Matriz[1][2].trim().length()>0  ){
 	     		 sql = "update TAF_REPORTE set nombre = '" + Matriz[1][2].trim().toUpperCase() + "'  where taf_path = '" +MatrizCarlos[x][1].trim() + "'";
 	     		 insert = operaBD.SQL_Ejecuta_Accion_Update_Insert_Delete(sql, 7, 0, 0);		    	 
 	     		 System.out.println("Es ok ?¡ : " + insert);
 	     		 System.out.println("");
     		  }else{
     			  //System.out.println("No se Rescato FECHA");
     			  //System.out.println("");
     		  }
    		  
    		  
    		  
     	  }else{
     		  String yh="";
     	  }
	      
	      
    	  if(caso.trim().equals("hojaConductor")){	
    		 sql = "update TAF_REPORTE set  hoja_anexo = " + Matriz[2][2].trim()+ " , " ;
    		 sql+= " hoja = " + Matriz[2][1].trim() + ", conductor=" + Matriz[2][3].trim() ;
    		 sql+= " where TAF_PATH = '" +MatrizCarlos[x][1].trim() + "'"; 
    		 insert = operaBD.SQL_Ejecuta_Accion_Update_Insert_Delete(sql, 7, 0, 0);		    	 
    		 System.out.println("Es ok ?¡ : " + insert);
    		 System.out.println("");
    	  }else{
    		  String yh="";
    	  }
    	  
    	  if(caso.trim().equals("ord")&&pasa.trim().equals("SI")){	
    		  
    		  if(Matriz[1][1].trim().length()>0  ){    			  
     		 sql = "update TAF_REPORTE set  ordinario = '" + Matriz[1][1].trim() + "'";     		 
     		 sql+= " where TAF_PATH = '" +MatrizCarlos[x][1].trim() + "'"; 
     		 insert = operaBD.SQL_Ejecuta_Accion_Update_Insert_Delete(sql, 7, 0, 0);		    	 
     		 System.out.println("Es ok ?¡ : " + insert);
     		 System.out.println("");
    		  }else{
    			  System.out.println("No se Rescato ORD");
    			  System.out.println("");
    		  }
     	  }else{
     		  String yh="";
     	  }
    		
    	  if(caso.trim().equals("folio")&&pasa.trim().equals("SI")){	
    		  
    		  if(Matriz[1][1].trim().length()>0  ){    			  
     		 sql = "update TAF_REPORTE set  folio = '" + Matriz[1][1].trim() + "'";     		 
     		 sql+= " where TAF_PATH = '" +MatrizCarlos[x][1].trim() + "'"; 
     		 insert = operaBD.SQL_Ejecuta_Accion_Update_Insert_Delete(sql, 7, 0, 0);		    	 
     		 System.out.println("Es ok ?¡ : " + insert);
     		 System.out.println("");
    		  }else{
    			  System.out.println("No se Rescato FOLIO");
    			  System.out.println("");
    		  }
     	  }else{
     		  String yh="";
     	  }


    	  if(caso.trim().equals("fecha")&&pasa.trim().equals("SI")){	
    		  
    		  if(Matriz[1][2].trim().length()>0  ){    			  
     		 sql = "update TAF_REPORTE set  fecha_solicitud = '" + Matriz[1][2].trim() + "'";     		 
     		 sql+= " where TAF_PATH = '" +MatrizCarlos[x][1].trim() + "'"; 
     		 insert = operaBD.SQL_Ejecuta_Accion_Update_Insert_Delete(sql, 7, 0, 0);		    	 
     		 System.out.println("Es ok ?¡ : " + insert);
     		 System.out.println("");
    		  }else{
    			  System.out.println("No se Rescato FECHA");
    			  System.out.println("");
    		  }
     	  }else{
     		  String yh="";
     	  }
    	  
    	  if(caso.trim().equals("totales")&&pasa.trim().equals("SI")){	
    		  
    		  if(Matriz[1][2].trim().length()>0  ){    			  
     		 sql = "update TAF_REPORTE set debe = " + Matriz[1][2].trim() + " , haber = " + Matriz[1][3].trim();     		 
     		 sql+= " where TAF_PATH = '" +MatrizCarlos[x][1].trim() + "'"; 
     		 insert = operaBD.SQL_Ejecuta_Accion_Update_Insert_Delete(sql, 7, 0, 0);		    	 
     		 System.out.println("Es ok ?¡ : " + insert);
     		 System.out.println("");
    		  }else{
    			  System.out.println("No se Rescato FECHA");
    			  System.out.println("");
    		  }
     	  }else{
     		  String yh="";
     	  }

    	  
    	  if(caso.trim().equals("rutnomb")&&pasa.trim().equals("SI")){	
    		  
    		  if(Matriz[1][2].trim().length()>0  ){    			  
     		 sql = "update TAF_REPORTE set rut_arc = " + Matriz[1][2].trim() + " , dv_arc = '" + Matriz[1][3].trim() + "'";     		 
     		 sql+= " where TAF_PATH = '" +MatrizCarlos[x][1].trim() + "'"; 
     		 insert = operaBD.SQL_Ejecuta_Accion_Update_Insert_Delete(sql, 7, 0, 0);	
     		 System.out.println("Es ok ?¡ : " + insert);
     		 
     		 sql = "update TAF_REPORTE set imponente = '" + Matriz[1][4].trim()+"' "  ;     		 
     		 sql+= " where TAF_PATH = '" +MatrizCarlos[x][1].trim() + "'"; 
     		 insert = operaBD.SQL_Ejecuta_Accion_Update_Insert_Delete(sql, 7, 0, 0);
     		 
     		 System.out.println("Es ok ?¡ : " + insert);
     		 System.out.println("");
    		  }else{
    			  System.out.println("No se Rescato FECHA");
    			  System.out.println("");
    		  }
     	  }else{
     		  String yh="";
     	  }

    	  
    	  
    	  
    	  
    	  
    	  
    	  try{
    		  int delay = 150; 							    					    	
    		  Thread.sleep( delay );							    	
    	 }catch(Exception e){}
	    	  
	    	  //sql = "update TAF_REPORTE set  hoja_anexo = '" + Matriz[1][2].trim().toUpperCase()+ "' where TAF_PATH = '" +MatrizCarlos[x][1].trim() + "'";
	    	  //sql = "update TAF_REPORTE set  hoja_anexo = " + Matriz[1][2].trim().toUpperCase()+ " where TAF_PATH = '" +MatrizCarlos[x][1].trim() + "'";
	    	  //System.out.println(sql);
	    	  //insert = operaBD.SQL_Ejecuta_Accion_Update_Insert_Delete(sql, 7, 0, 0);		    	 
	    
	      insert = "OK";			      
	      
	      operaBD = null;
	      Matriz = null;
	      if(MatrizCarlos[x+1][1].trim().equals("FIN")){x=9999999;}
	}
	
	System.out.println("Fin");
}catch(Exception e){System.out.println("Error Archivo : " + fileName + " - " + e);}	
}





























public static void ProcesoDetalle(){
	int h = 0;	
	String fileName = "" , insert = "";
	 
	DevuelveDetalleTAF_Detall back = new DevuelveDetalleTAF_Detall();
	try{
	String MatrizCarlos[][] = new String[15000][500];	
	String Matriz[][] = new String[15000][500];
	
	
		
	 String sql = " select taf_path, hoja_anexo"; 
	 sql+= " from taf_reporte where analisar = 'X'";
	 
	 
	
	MatrizCarlos = back.CasosParaAnalizar(sql);
	
	
	 sql = " select count(taf_path) "; 
	 sql+= " from taf_reporte where analisar = 'X'";
	
	
	
	int valor = back.cuentaCasos(sql);
	

																									

	int ok = 0,x = 0;
	String cumple = "",detalle = "",error = "" ;
	
	for(x=1; x<valor; x++){	
		
		
		h = 0;
		Bd operaBD = new Bd();	
		
		fileName ="C:\\Opt\\Xls\\" + MatrizCarlos[x][1].trim();
		//fileName ="C:\\Opt\\Xls\\NuevoMaterial\\CORREGIDOS_2012\\34877815.xls";
		//fileName ="C:\\Opt\\Xls\\NuevoMaterial\\9688542-3.xls";
		//fileName ="C:\\Opt\\Xls\\NuevoMaterial\\CORREGIDOS_2012\\4.121.459-7 PLANVITAL.xls";
		
		
		int hoja =  Integer.parseInt(MatrizCarlos[x][2].trim());
		//hoja = 1;
		//hoja=2;	
		if(hoja>0){
		//fileName ="C:\\Opt\\Xls\\2010\\48843352(E).xls" ;
			
		
		System.out.println("");
		System.out.println("ES : " + x + " de " + valor + "  --->  " + fileName);		
		
		Matriz = back.DevuelveDetalle(fileName , hoja);	
		operaBD = null;
		}else{
			System.out.println("");
			System.out.println("ES : hoja = 0  --->  " + fileName);			
		}
		
		
	}

}catch(Exception e){System.out.println("Error Proceso Main - ProcesoDetalle : " + fileName + " - " + e);}	
	
}


}



				