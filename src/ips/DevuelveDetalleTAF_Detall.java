package ips;

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

public class DevuelveDetalleTAF_Detall {
	
	public String[][] DevuelveDetalle(String fileName, int hoja) {
		
		
		String matrizX[][] = new String[2000][300];
		String matrizX2[][] = new String[2000][300];
		
		matrizX = LlenaMAtrizXReadExcel(fileName, hoja);
		
		
		int filasX = Integer.parseInt(matrizX[1995][298]);
		int colX = Integer.parseInt(matrizX[1995][299]); 
		
		/*
		for(int j=0;j<=filasX;j++){
	    	for(int k=0;k<=colX;k++){		    		
	    		if(matrizX[j][k]==null){}else{
	    			System.out.println("matrizx" + "[" + j + "]" + "[" + k + "]" +" = " + matrizX[j][k]);
	    		}
	    	}
	    }
		*/
		
				
		// Obtener Cantidad de Bloques
		int  numeroBloques = 0;
		for (int fil = 1; fil <= filasX; fil++) {
			for (int col = 0; col <= colX; col++) {		
				if(matrizX[fil][col]==null){}else{
	    			//System.out.println("matrizx" + "[" + fil + "]" + "[" + col + "]" +" = " + matrizX[fil][col]);
	    			int posibilidades = matrizX[fil][col].trim().indexOf("MONTO R");
					if (posibilidades== -1 ){}else{numeroBloques++;}	
	    		}					
				}
			}
		
	System.out.println("Cantidad de Bloques es : " + numeroBloques);
	
	
	String nombreEmpleador = "";
	String rutEmpleador = "";
	String montoRenta = "";
	
	int ghf = 0;
	int gfdghf = 0;
	for(int g = 1;g<=numeroBloques;g++){
		
		if(g==24){
			String gdg = "";
		}
		
		for (int fil = 1; fil <= filasX; fil++) {
			for (int col = 0; col <= 300; col++) {		
				if (matrizX[fil][col] == null) {
					ghf++;
					if(ghf==5){
						col = 400;	ghf=0;
					}
					
				} else {
					if (!matrizX[fil][col].trim().equals("") && !matrizX[fil][col].trim().equals("null")) {									
						
						// Nombre Empleador
						try{
						int posibilidades = matrizX[fil][col].trim().indexOf("NOMBRE EM");
						if (posibilidades== -1 ){
							posibilidades = matrizX[fil][col].trim().indexOf("MONTO REN");
							if (posibilidades== -1 ){
							}else{
								nombreEmpleador = "SIN NOMBRE REGISTRADO";
								matrizX2 = ObtengoDetalleMontosLlenoMAtrizInsert(matrizX,fil-2,col, nombreEmpleador, rutEmpleador, fileName , filasX , colX);
								matrizX=null;
								matrizX = matrizX2;
								
								col = 9999;
								fil = 999;
								
							}
						}else{
							
								String ok = "0";
								int col_ = 1;
								int loops = 0;
								
								if(matrizX[fil][col+col_]==null){nombreEmpleador="";loops++;;}else{
									nombreEmpleador = matrizX[fil][col+col_].trim();	
									loops++;
								}
																
								// Busca Nombre Empleador y rut Empleador
								while(ok.trim().equals("0")){
									
									try{
										if(!nombreEmpleador.trim().equals("")){
											if(nombreEmpleador.trim().length()>3){
												ok = "1";
												
												matrizX2 = ObtengoDetalleMontosLlenoMAtrizInsert(matrizX,fil,col, nombreEmpleador, rutEmpleador, fileName , filasX , colX);
												matrizX=null;
												matrizX = matrizX2;
												
												col = 9999;
												fil = 999;
												loops=0;
											}else{
												col_++;
												if(matrizX[fil][col+col_]==null){nombreEmpleador="";}else{nombreEmpleador = matrizX[fil][col+col_].trim();}												
												if(matrizX[fil+1][col+col_]==null){rutEmpleador="";}else{rutEmpleador = matrizX[fil+1][col+col_].trim();}												
											}
										}else{
											col_++;										
											if(matrizX[fil][col+col_]==null){nombreEmpleador="";}else{nombreEmpleador = matrizX[fil][col+col_].trim();}											
											if(matrizX[fil+1][col+col_]==null){rutEmpleador="";}else{rutEmpleador = matrizX[fil+1][col+col_].trim();}
										}
									loops++;
									
									if(loops>4){
										nombreEmpleador = "SIN NOMBRE REGISTRADO";
									}
									
									}catch(Exception e){
										System.out.println("Error Ciclo Intreno Nombre EMPL : " + e);ok="1";
									}
									
								}
								/////
							
							}						
						}catch(Exception e){
							System.out.println("Error Nombre Em Ciclo Externo  : " + e);
							}
						//////////////////////////////////////////////////////////////////////////////////////////////////////
					
						
					}else{
						gfdghf++;
						if(gfdghf==10){
							col = 400;	gfdghf=0;
						}
						
					}	
				}
			}	
	
		}
		
	}
	
		return matrizX;
	}

	public int cuentaCasos(String sql) {
		String Retorno[][] = new String[2][2];
		Bd operaBD = new Bd();		
		Retorno = operaBD.SQL_Dev_Matriz_Env_Sql(sql, 1, 7);
		int valor = Integer.valueOf(Retorno[1][1].trim());
		Retorno = null;
		return valor;
	}

	public String[][] CasosParaAnalizar(String sql) {
		String MatrizCarlos[][] = new String[15000][10];	
		Bd operaBD = new Bd();
		MatrizCarlos = operaBD.SQL_Dev_Matriz_Env_Sql(sql, 2, 7);
		operaBD = null;
		return MatrizCarlos;
	}
	
	public String[][] CasosParaAnalizarConduct(String sql) {
		String MatrizCarlos[][] = new String[15000][10];	
		Bd operaBD = new Bd();
		MatrizCarlos = operaBD.SQL_Dev_Matriz_Env_Sql(sql, 1, 7);
		operaBD = null;
		return MatrizCarlos;
	}

	public String[][] ObtengoDetalleMontosLlenoMAtrizInsert(String matrizX[][] , int fil, int col, String nombreEmpleador , String rutEmpleador, String fileName, int filasX, int colX){
		
		String digitoRut = "";
		try{
			int largoRut = rutEmpleador.length();
			if(largoRut>1){
				rutEmpleador = rutEmpleador.replace(".", "");
				rutEmpleador = rutEmpleador.replace("/", "");
				rutEmpleador = rutEmpleador.replace("-", "");
				rutEmpleador = rutEmpleador.replace(" ", "");
				rutEmpleador = rutEmpleador.replace(",", "");
				largoRut = rutEmpleador.length();
				digitoRut = rutEmpleador.substring(largoRut-1, largoRut);
				rutEmpleador = rutEmpleador.substring(0, largoRut-1);
			}else{rutEmpleador="";digitoRut="";}
		}catch(Exception e){digitoRut = "";}
		
		int marcaColumna = 0;
		String matrizX_ANT[][] = new String[2000][300];
		for(int i=fil;i<=100;i++){
			if (matrizX[i][col]==null) {}else{
				int posibilidades = matrizX[i][col].trim().indexOf("MONTO");
				
				if(posibilidades==-1){}else{
					i++;
					String salir = "NO";
					int breaker = 0;
					while(salir.trim().equals("NO")){
						i++;
						breaker++;
						if(breaker>15){
							i=9999;
							marcaColumna = col+7;
							matrizX_ANT = restructuraMatriz(matrizX, filasX, colX, marcaColumna);
							matrizX = null;
							matrizX = matrizX_ANT;
							break;
						}
						if(matrizX[i][col]==null){
							
						}else{
							
							if(!matrizX[i][col].trim().equals("")){
								
								String continua = "SI";
								while (continua.trim().equals("SI")){
									String montoRenta ="";
									String mes = "";
									String ano = "";
									String dif = "";
									String reten = "";
									String montoTraspasar = "";
									
									if(matrizX[i][col]==null){}else{montoRenta = matrizX[i][col].trim();}									
									if(matrizX[i][col+3]==null){}else{mes = matrizX[i][col+3].trim();}
									if(matrizX[i][col+4]==null){}else{ano = matrizX[i][col+4].trim();}
									if(matrizX[i][col+5]==null){}else{dif = matrizX[i][col+5].trim();}
									if(matrizX[i][col+6]==null){}else{reten = matrizX[i][col+6].trim();}
									
																		
									if(mes.trim().toUpperCase().equals("ENERO")){mes="01";}
									if(mes.trim().toUpperCase().equals("FEBRERO")){mes="02";}
									if(mes.trim().toUpperCase().equals("MARZO")){mes="03";}
									if(mes.trim().toUpperCase().equals("ABRIL")){mes="04";}
									if(mes.trim().toUpperCase().equals("MAYO")){mes="05";}
									if(mes.trim().toUpperCase().equals("JUNIO")){mes="06";}
									if(mes.trim().toUpperCase().equals("JULIO")){mes="07";}
									if(mes.trim().toUpperCase().equals("AGOSTO")){mes="08";}
									if(mes.trim().toUpperCase().equals("SEPTIEMBRE")){mes="09";}
									if(mes.trim().toUpperCase().equals("OCTUBRE")){mes="10";}
									if(mes.trim().toUpperCase().equals("NOVIEMBRE")){mes="11";}
									if(mes.trim().toUpperCase().equals("DICIEMBRE")){mes="12";}
									
									
									try{
									if(matrizX[i][col+7]==null){
										
									}else{
										montoTraspasar = matrizX[i][col+7].trim();
										if(montoTraspasar.trim().equals("")){
											montoTraspasar = matrizX[i][col+6].trim();
											try{												
												String j = montoTraspasar.replace(".", "0");
												int testing = Integer.valueOf(j) + 6;
											}catch(Exception e){montoTraspasar="";}											
										}else{
											String hjk = "";
											try{
												
												String j = montoTraspasar.replace(".", "0");
												int testing = Integer.valueOf(j) + 6;
												
											}catch(Exception e){hjk="Error";System.out.println("Error:" + e);}		
											
											if(hjk.trim().equals("Error")){
												montoTraspasar = matrizX[i][col+6].trim();
											}
										}
									}	
								}catch(Exception e){
									continua="NO";
									}
									
								 
									String tot = "";
									int xposx = montoRenta.trim().indexOf("TOTA");if (xposx== -1 ){}else{tot = "SI";}
									xposx = mes.trim().indexOf("TOTA");if (xposx== -1 ){}else{tot = "SI";}	
									xposx = ano.trim().indexOf("TOTA");if (xposx== -1 ){}else{tot = "SI";}	
									xposx = dif.trim().indexOf("TOTA");if (xposx== -1 ){}else{tot = "SI";}	
									xposx = reten.trim().indexOf("TOTA");if (xposx== -1 ){}else{tot = "SI";}	
									xposx = montoTraspasar.trim().indexOf("TOTA");if (xposx== -1 ){}else{tot = "SI";}
									
									if(tot.trim().equals("SI")){montoRenta="0";mes="0";ano="0";dif="0";reten="0";montoTraspasar="0";}
								
									marcaColumna = col+7;
									//if( (montoRenta.trim().length()>0) && (mes.trim().length()>0)&& (ano.trim().length()>0)){
									if( (montoRenta.trim().length()>0) ){	
									}else{			
										
										
										matrizX_ANT = restructuraMatriz(matrizX, filasX, colX, marcaColumna);
										matrizX = null;
										matrizX = matrizX_ANT;
										
										i=999;
										break;
									}
										
									
									long MontoRentaL=0 , mesL=0, anoL=0, difL=0, retenL=0, montoTraspasarL=0;
									double montoRentaD=0, mesD=0, anoD=0, difD=0, retenD=0, montoTraspasarD=0;
									
									try{
										if(montoRenta.trim().length()>0){
											try{
											montoRentaD = Double.valueOf(montoRenta).doubleValue();
											MontoRentaL =   Math.round(montoRentaD);
											}catch(Exception e){MontoRentaL=0;}
										}
										if(mes.trim().length()>0){
											
											try{
											mesD = Double.valueOf(mes).doubleValue();
											mesL =   Math.round(mesD);
											}catch(Exception e){mesL=0;}
										}
										if(ano.trim().length()>0){
											try{
											anoD = Double.valueOf(ano).doubleValue();
											anoL =   Math.round(anoD);
											}catch(Exception e){anoD=0;}
										}
										if(dif.trim().length()>0){
											try{
											dif = dif.replace(",", ".");											
											difD = Double.valueOf(dif).doubleValue();
										}catch(Exception e){difD=0;}
											//difL =   Math.round(difD);
										}
										if(reten.trim().length()>0){
											try{
											reten = reten.replace(",", ".");	
											retenD = Double.valueOf(reten).doubleValue();
										}catch(Exception e){retenD=0;}
											//retenL =   Math.round(retenD);
										}
										if(montoTraspasar.trim().length()>0){
											try{
											montoTraspasarD = Double.valueOf(montoTraspasar).doubleValue();
											montoTraspasarL =   Math.round(montoTraspasarD);
										}catch(Exception e){montoTraspasarL=0;}
										}
										
										/*
										if(difD>0){}else{
											if(retenD>0){difD=retenD;}
										}
										*/
										
										
									}catch(Exception e){
										System.out.println("No es valor " + e);
										}
									
									
									nombreEmpleador = nombreEmpleador.replace(":", "");									
									nombreEmpleador = nombreEmpleador.replace(",", "");
									nombreEmpleador = nombreEmpleador.replace(";", "");
									rutEmpleador = rutEmpleador.replace(":", "");									
									rutEmpleador = rutEmpleador.replace(",", "");
									rutEmpleador = rutEmpleador.replace(";", "");
									digitoRut = digitoRut.replace(":", "");									
									digitoRut = digitoRut.replace(",", "");
									digitoRut = digitoRut.replace(";", "");
									
									try{
										int gfds = 0;
										gfds = Integer.parseInt(rutEmpleador);
									}						
									catch(Exception e){rutEmpleador = "";digitoRut="";}
									
									
									String w = "0";
									String sql = " insert into TAF_REPORTE_DETALLE (";	
									if(fileName.trim().length()>0){sql+= " taf_path ";w="1";}								
									if(rutEmpleador.trim().length()>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= " RUT_EMP ";w="1";}									
									if(digitoRut.trim().length()>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= " DV_EMP ";w="1";}								
									if(nombreEmpleador.trim().length()>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= " EMPLEADOR ";w="1";}								
									if(MontoRentaL>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= " RENTA ";w="1";}
									if(anoL>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= " ANIO ";w="1";}	
									if(mesL>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= " MES ";w="1";}	
									if(difD>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= " TASA ";w="1";}	
									if(montoTraspasarL>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= " MONTO ";w="1";}
									sql+=" , nro) values (";
									if(fileName.trim().length()>0){
										int lrg = fileName.length();
										String fileName2 = fileName.substring(25, lrg);
										sql+= "'" + fileName2.trim() + "'";w="1";
										
									}	
									if(rutEmpleador.trim().length()>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= "" + rutEmpleador+ "";w="1";}	
									if(digitoRut.trim().length()>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= "'" + digitoRut+ "'";w="1";}
									if(nombreEmpleador.trim().length()>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= "'" +nombreEmpleador+ "'";w="1";}
									if(MontoRentaL>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= MontoRentaL;w="1";}	
									if(anoL>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= anoL;w="1";}	
									if(mesL>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= "'" + mesL+ "'";w="1";}	
									if(difD>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= difD;w="1";}	
									if(montoTraspasarL>0){if(w.trim().equals("1")){sql+=" , ";w="0";}sql+= montoTraspasarL;w="1";}
									sql+=" , " + i;
									sql+=" )";
									
									Bd operaBD = new Bd();
									String insert = "";
									
									
									if(MontoRentaL>0 ){										
										insert = operaBD.SQL_Ejecuta_Accion_Update_Insert_Delete(sql, 7, 0, 0);
										//System.out.println("SQL: " + sql);
										//System.out.println("");
										//System.out.println(sql);
									}
							    	operaBD = null;
								
							    	try{
							    	int delay = 10; 							    					    	
							    	Thread.sleep( delay );							    	
							    	}catch(Exception e){}
							    	
									//i=999;
									//System.out.println("Resultado Insert de : " +fileName + " ---> " +  insert);
									//System.out.println("Sql :" + sql);
									//System.out.println("");
									salir = "SI";
								
									
									i++;
									
									
									}
									
								}
						
						}	
							
							
						}
						i++;
					}
					
					
					
					
				}
			}
		
	
	//System.out.println(fileName + " - " + "MontoRenta : " + montoRenta);
		
		
		
	return matrizX;
	}
	
	public String [][] restructuraMatriz (String MatrizX[][], int filasX, int colX, int ending){
		String matrizX[][] = new String[2000][300];
		String matrizX_ANT[][] = new String[2000][300];
		
		
		for(int fil=0;fil<=filasX;fil++){	
			for(int col=ending+1;col<=colX;col++){
				if (MatrizX[fil][col]==null){
						//System.out.println("<<>> - MatrizX[" + fil + "]" + "[" + col + "]= NADA"  ); 
						//matrizX_ANT[fil][col-ending+1] = null;						
						//System.out.println("XXXXX - MatrizXANT[" + fil + "]" + "[" + (col-ending+1) + "]= NADA"  );						
					}else{
						//System.out.println("<<>> - MatrizX[" + fil + "]" + "[" + col + "]=" + MatrizX[fil][col] ); 		
						matrizX_ANT[fil][col-ending-1] = MatrizX[fil][col];						
						//System.out.println("XXXXX - MatrizXANT[" + fil + "]" + "[" + (col-ending-1) + "]= " +  matrizX_ANT[fil][col-ending-1]  );						
				}				
			}			
		}
		
		/*
		for(int i=0;i<=39;i++){
			for(int j=1;j<=32;j++){
				if( matrizX_ANT[i][j]==null){}else{
				System.out.println("!!!!!!!! - MatrizXANT[" + i + "]" + "[" + j + "]= " +  matrizX_ANT[i][j]  );
				}
			}
		}
		*/
		
		matrizX = null;
		matrizX = matrizX_ANT;
		
		return matrizX;
	}
	
	public String[][] LlenaMAtrizXReadExcel (String fileName, int hoja){
		String matrizX[][] = new String[2000][300];
		String ficheroEntrada = fileName;		   
		try{
	    HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(ficheroEntrada));
	    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
	    
	    HSSFSheet sheet = workbook.getSheetAt(hoja-1);
	    HSSFRow row = null;
	    HSSFCell hssfCell = null;
	    
	    CellValue cellValue = evaluator.evaluate(hssfCell);
	    
	    int filasX = 0;
	    int colX=0;
	    for ( int i=0, z=sheet.getLastRowNum(); i<z; i++ ){
	    	row = sheet.getRow(i);
	    	if (row != null) {		    		
	    		for (int ii=0, zz=row.getLastCellNum(); ii<zz; ii++){
	    			
	    			if(ii>colX){colX=ii;}
	    			
	    			hssfCell = row.getCell((short) ii);
	        		        
	    			if (hssfCell != null) {
	    				
	    				// Cell type (0) / numeric
	    				// Cell type (1) /( string)
	    				// Cell type (2) / formula
	    				// Cell type (3) / blank
	    				// Cell type (4) / boolean
	    				// Cell type (5) / Error
	    				
	    				if(hssfCell.getCellType()==0){
	    					double g = hssfCell.getNumericCellValue();		    					
	    					//long g2 = Math.round(g); 
	    					matrizX[i][ii] =  Double.toString(g);
	    					//System.out.println("DOUBLE   - linea " + i + " col " + ii + " " + g);
	    				}
	    				
	    				if(hssfCell.getCellType()==1){
	    					String  g = hssfCell.getStringCellValue();		    					
	    					matrizX[i][ii] = g; 
	    					//System.out.println("STRING   - linea " + i + " col " + ii + " " + g);
	    				}
	    				
	    				if(hssfCell.getCellType()==2){
	    					HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(workbook);
	    					CellValue cv = null;
	    					cv = fe.evaluate(hssfCell);
	    					
	    					if(cv.getCellType()==2){
	    						String g = cv.getStringValue();		    						
	    						matrizX[i][ii] = g; 
	    						//System.out.println("STRING F - linea " + i + " col " + ii + " " + g);
	    					}
	    					if(cv.getCellType()==1){
	    						String g = cv.getStringValue();		    						
	    						matrizX[i][ii] = g; 
	    						//System.out.println("STRING F - linea " + i + " col " + ii + " " + g);
	    					}
	    					if(cv.getCellType()==0){
	    						double g = cv.getNumberValue();		    						
	    						long g2 = Math.round(g); 
		    					matrizX[i][ii] = String.valueOf(g2); 
		    					//System.out.println("NUMBER F - linea " + i + " col " + ii + " " + g);
	    					}		    					
	    				}
	    				
	    			}else{matrizX[i][ii] = "";}		    			
	    		}
	    	}
	    	filasX++;
	    }
	  
	
	
	matrizX[1995][298] = String.valueOf(filasX); 
	matrizX[1995][299] = String.valueOf(colX); 
	
}catch(Exception e) {System.out.println("Excepción en ReadXL : " + e );}

		return matrizX;
	}
	
}
