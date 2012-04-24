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

public class DevuelveDetalleTAF_Conduc {
	
	
		public String[][] DevuelveConductorHoja(String fileName) {
			String matriz[][] = new String[2000][100];
			
			matriz[2][4]="0";
			matriz[2][2]="0";
			matriz[2][1]="0";
			matriz[2][3]="0";
			matriz[1][1] = ""; matriz[1][2] = "";
			int hojasTotal = 0;
			List cellDataList = new ArrayList();
			try {
				FileInputStream fileInputStream = new FileInputStream(fileName);
				POIFSFileSystem fsFileSystem = new POIFSFileSystem(fileInputStream);
				HSSFWorkbook workBook = new HSSFWorkbook(fsFileSystem);
				FormulaEvaluator evaluator = workBook.getCreationHelper().createFormulaEvaluator();
				hojasTotal = workBook.getNumberOfSheets();
				int numeroHojaConductor = 0;
				System.out.println("Archivo : " + fileName);
				for(int i=0;i<hojasTotal;i++){
					HSSFSheet hssfSheet = workBook.getSheetAt(i);
					Iterator rowIterator = hssfSheet.rowIterator();
					while (rowIterator.hasNext()){
						HSSFRow hssfRow = (HSSFRow) rowIterator.next();
						Iterator iterator = hssfRow.cellIterator();
						while (iterator.hasNext()){
							List cellTempList = new ArrayList();
							HSSFCell hssfCell = (HSSFCell) iterator.next();
							CellValue cellValue = evaluator.evaluate(hssfCell);
							String stringCellValue = hssfCell.toString();
							if(stringCellValue.trim().equals("")){}else{
								if(hssfCell.getCellType()==2){
									HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(workBook);
									CellValue cv = null;
									try {
										cv = fe.evaluate(hssfCell);
										if(cv.getCellType()==4){
											hssfCell.setCellType(1);
											hssfCell.setCellValue(cv.getBooleanValue());
										}
										if(cv.getCellType()==2){
											hssfCell.setCellType(1);
											hssfCell.setCellValue(cv.getStringValue());
										}
										if(cv.getCellType()==1){
											hssfCell.setCellType(1);
											hssfCell.setCellValue(cv.getStringValue());
										}
										if(cv.getCellType()==0){
											hssfCell.setCellType(0);
											hssfCell.setCellValue(cv.getNumberValue());
										}
									} catch (Exception e) {matriz[1][1]= "";}
								}
							
								int posibilidades1 = stringCellValue.indexOf("ANTECEDENTES");
								if (posibilidades1== -1 ){}else{
									matriz[1][1]+="1";
								}                     
					
								int posibilidades2 = stringCellValue.indexOf("INFORME");
								if (posibilidades2== -1 ){}else{
									matriz[1][1]+="2";
								}
					
								int posibilidades3 = stringCellValue.indexOf("OBSERVACIONES");
								if (posibilidades3== -1 ){}else{
									matriz[1][1]+="3";
								}
								
								int posibilidades4 = stringCellValue.indexOf("ANEXO LIQ");
								if (posibilidades4== -1 ){}else{									
									matriz[2][2] = String.valueOf(i+1);									
								}
					
								
								
								
								
								if(matriz[1][1].trim().equals("123")){									
									matriz[2][1] = String.valueOf(i+1);
									matriz[2][3] = "1";
									matriz[2][4] = "OK";
									matriz[1][1] = "fdswe";
								}
								
							
																								
								cellTempList.add(hssfCell);
								if(cellTempList.isEmpty()){}else{cellDataList.add(cellTempList);}
							}
						}
					}
					hssfSheet = null;
					rowIterator = null;
				} // End For
				
				
				
				fileInputStream = null;
				fsFileSystem = null;
				workBook =null;
				evaluator = null;
				}catch (Exception e) {
					System.out.println("Error en Busquedas de Hojas : " + fileName + " - " + e);
					}
			return matriz;
		}
	
		
		
	
		
		
		
		public String[][] DevuelveConductorOrd(String fileName, int hoja) {
			String matriz[][] = new String[2000][100];
			matriz[1][1] = "";
			matriz[1][2] = "";
			int hojasTotal = 0;

			List cellDataList = new ArrayList();



			try {

			FileInputStream fileInputStream = new FileInputStream(fileName);
			POIFSFileSystem fsFileSystem = new POIFSFileSystem(fileInputStream);
			HSSFWorkbook workBook = new HSSFWorkbook(fsFileSystem);
			FormulaEvaluator evaluator = workBook.getCreationHelper().createFormulaEvaluator();

			int i = hoja-1;
			hojasTotal = i;
			//for(int i=0;i<hojasTotal;i++){

			HSSFSheet hssfSheet = workBook.getSheetAt(i);
			Iterator rowIterator = hssfSheet.rowIterator();
			while (rowIterator.hasNext()){
			HSSFRow hssfRow = (HSSFRow) rowIterator.next();
			Iterator iterator = hssfRow.cellIterator();



			while (iterator.hasNext()){
			List cellTempList = new ArrayList();
			HSSFCell hssfCell = (HSSFCell) iterator.next();
			CellValue cellValue = evaluator.evaluate(hssfCell);
			String stringCellValue = hssfCell.toString();

			if(stringCellValue.trim().equals("")){}else{
			if(hssfCell.getCellType()==2){
			HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(workBook);
			CellValue cv = null;

			try {
			cv = fe.evaluate(hssfCell);

			if(cv.getCellType()==4){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getBooleanValue());
			}

			if(cv.getCellType()==2){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getStringValue());
			}

			if(cv.getCellType()==1){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getStringValue());
			}

			if(cv.getCellType()==0){
			hssfCell.setCellType(0);
			hssfCell.setCellValue(cv.getNumberValue());
			}

			} catch (Exception e) {matriz[1][1]= "";}
			}

			cellTempList.add(hssfCell);
			if(cellTempList.isEmpty()){}else{cellDataList.add(cellTempList);}

			}

			}



			}
			hssfSheet = null;
			rowIterator = null;

			//} // End For

			fileInputStream = null;
			fsFileSystem = null;
			workBook =null;
			evaluator = null;
			}catch (Exception e) {System.out.println("Error en 1 Ciclo Devuelve : " + fileName + "-" + e);}

			String ord ="",folio = "",total = "";








			try{

			int posibilidades = 0;
			cellDataList = cellDataList;
			HSSFCell rut1 = null;
			String matrizX[][] = new String[2000][100];

			int lineas = cellDataList.size(); // Cantidad de lineas
			for (int i = 0; i < cellDataList.size(); i++){
			List cellTempList = (List) cellDataList.get(i);

			for (int j = 0; j < cellTempList.size(); j++){
			HSSFCell hssfCell = (HSSFCell) cellTempList.get(j);
			String stringCellValue = hssfCell.toString();
			stringCellValue=stringCellValue.toUpperCase();


			

			posibilidades = stringCellValue.indexOf("ORD");
			if (posibilidades== -1 ){matriz[1][1] = "";}else{
				List cellTempList2 = (List) cellDataList.get(i+1);
				rut1 = (HSSFCell) cellTempList2.get(j);
				matriz[1][1]=rut1.toString();
				System.out.println("Es : " + rut1.toString());
				j=999;
				i=999;

			}
			}
			}
			String hghg = "";
			}catch(Exception e){System.out.println("Error en Busqueda de patrones: " + e + " en archivo : " + fileName);}
			return matriz;
		}
	
	
		
		
		
		
		
		
		
		
		
		public String[][] DevuelveConductorFechaFak(String fileName, int hoja) {
			String matriz[][] = new String[2000][100];
			matriz[1][1] = "";
			matriz[1][2] = "";
			int hojasTotal = 0;

			List cellDataList = new ArrayList();



			try {

			FileInputStream fileInputStream = new FileInputStream(fileName);
			POIFSFileSystem fsFileSystem = new POIFSFileSystem(fileInputStream);
			HSSFWorkbook workBook = new HSSFWorkbook(fsFileSystem);
			FormulaEvaluator evaluator = workBook.getCreationHelper().createFormulaEvaluator();

			int i = hoja-1;
			hojasTotal = i;
			//for(int i=0;i<hojasTotal;i++){

			HSSFSheet hssfSheet = workBook.getSheetAt(i);
			Iterator rowIterator = hssfSheet.rowIterator();
			while (rowIterator.hasNext()){
			HSSFRow hssfRow = (HSSFRow) rowIterator.next();
			Iterator iterator = hssfRow.cellIterator();



			while (iterator.hasNext()){
			List cellTempList = new ArrayList();
			HSSFCell hssfCell = (HSSFCell) iterator.next();
			CellValue cellValue = evaluator.evaluate(hssfCell);
			String stringCellValue = hssfCell.toString();

			if(stringCellValue.trim().equals("")){}else{
			if(hssfCell.getCellType()==2){
			HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(workBook);
			CellValue cv = null;

			try {
			cv = fe.evaluate(hssfCell);

			if(cv.getCellType()==4){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getBooleanValue());
			}

			if(cv.getCellType()==2){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getStringValue());
			}

			if(cv.getCellType()==1){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getStringValue());
			}

			if(cv.getCellType()==0){
			hssfCell.setCellType(0);
			hssfCell.setCellValue(cv.getNumberValue());
			}

			} catch (Exception e) {matriz[1][1]= "";}
			}

			cellTempList.add(hssfCell);
			if(cellTempList.isEmpty()){}else{cellDataList.add(cellTempList);}

			}

			}



			}
			hssfSheet = null;
			rowIterator = null;

			//} // End For

			fileInputStream = null;
			fsFileSystem = null;
			workBook =null;
			evaluator = null;
			}catch (Exception e) {System.out.println("Error en 1 Ciclo Devuelve : " + fileName + "-" + e);}

			String ord ="",folio = "",total = "";








			try{

			int posibilidades = 0;
			cellDataList = cellDataList;
			HSSFCell rut1 = null;
			String matrizX[][] = new String[2000][100];

			int lineas = cellDataList.size(); // Cantidad de lineas
			String mat = "";
			String afp = "";
			for (int i = 0; i < cellDataList.size(); i++){
			List cellTempList = (List) cellDataList.get(i);

			
			
			for (int j = 0; j < cellTempList.size(); j++){
			HSSFCell hssfCell = (HSSFCell) cellTempList.get(j);
			String stringCellValue = hssfCell.toString();
			stringCellValue=stringCellValue.toUpperCase();
			
			if(mat.trim().equals("OK") && afp.trim().equals("OK")){
				j=999;
				i=999;
			}
			
			if(mat.trim().equals("OK")){}else{
				posibilidades = stringCellValue.indexOf("MAT");
				if (posibilidades== -1 ){}else{
					
					String cadenaCaracter = "";
					List cellTempList2;
					try{
						cellTempList2 = (List) cellDataList.get(i+2);				
						rut1 = (HSSFCell) cellTempList2.get(j);
						cadenaCaracter = rut1.toString();				
					}catch(Exception e){}
					
					try{
						cellTempList2 = (List) cellDataList.get(i+3);
						rut1 = (HSSFCell) cellTempList2.get(j);
						cadenaCaracter+= " - " + rut1.toString();				
					}catch(Exception e){}
					
										
					//System.out.println("Archivo  : " + fileName + "  -   " +  rut1.toString());
					String dia = "";
					String mes = "";
					String ano = "";
					String cadena = rut1.toString().trim();   
					int largo = cadena.length();
					if(largo==15){
						dia = stringCellValue.substring(10,12);
						mes = stringCellValue.substring(13,15);
						ano = "20" + stringCellValue.substring(16,18);
						//System.out.println("Dia " + dia);
						//System.out.println("Mes " + mes);
						//System.out.println("Ano " +ano);
						String fecha = dia+"/"+mes+"/"+ano;
						matriz[1][1]=fecha;	
					}

					if (largo==11){
						dia = cadena.substring(0,2);
						mes = cadena.substring(3,6);
						ano = cadena.substring(7,11);
						mes=mes.toUpperCase();
						if(mes.trim().equals("ENE")){mes="01";}
						if(mes.trim().equals("FEB")){mes="02";}
						if(mes.trim().equals("MAR")){mes="03";}
						if(mes.trim().equals("ABR")){mes="04";}
						if(mes.trim().equals("MAY")){mes="05";}
						if(mes.trim().equals("JUN")){mes="06";}
						if(mes.trim().equals("JUL")){mes="07";}
						if(mes.trim().equals("AGO")){mes="08";}
						if(mes.trim().equals("SEP")){mes="09";}
						if(mes.trim().equals("OCT")){mes="10";}
						if(mes.trim().equals("NOV")){mes="11";}
						if(mes.trim().equals("DIC")){mes="12";}
						String fecha = dia+"/"+mes+"/"+ano;
						//System.out.println("Es Valido : " + cadena + " MES : " + mes + " FIN :  " + fecha);
						matriz[1][1]=fecha;
					}else{
						//System.out.println("NO Es Valido : " + cadena);
					}
					
					if (largo==10){
						dia = cadena.substring(0,2);
						mes = cadena.substring(3,5);
						ano = cadena.substring(6,10);
						mes=mes.toUpperCase();
						if(mes.trim().equals("ENE")){mes="01";}
						if(mes.trim().equals("FEB")){mes="02";}
						if(mes.trim().equals("MAR")){mes="03";}
						if(mes.trim().equals("ABR")){mes="04";}
						if(mes.trim().equals("MAY")){mes="05";}
						if(mes.trim().equals("JUN")){mes="06";}
						if(mes.trim().equals("JUL")){mes="07";}
						if(mes.trim().equals("AGO")){mes="08";}
						if(mes.trim().equals("SEP")){mes="09";}
						if(mes.trim().equals("OCT")){mes="10";}
						if(mes.trim().equals("NOV")){mes="11";}
						if(mes.trim().equals("DIC")){mes="12";}
						String fecha = dia+"/"+mes+"/"+ano;
						//System.out.println("Es Valido : " + cadena + " MES : " + mes + " FIN :  " + fecha);
						matriz[1][1]=fecha;	
					}else{
						//System.out.println("NO Es Valido : " + cadena);
					}
					
					
					
					
					
					
					 if(matriz[1][1].trim().length()>0  ){}else{
					int pes = cadenaCaracter.indexOf("JEFE SECCION");
					int pes2 = cadenaCaracter.indexOf("DE :");
					
					if(pes==-1 && pes2==-1){
						System.out.println(cadenaCaracter);
						matriz[1][1] = cadenaCaracter;
					}
					}
					
					
					mat = "OK";
						
				}
			}
			
			
			
			if(afp.trim().equals("OK")){}else{
				posibilidades = stringCellValue.indexOf("ADMINISTRAD");
				if (posibilidades== -1 ){}else{
					
					String cadenaCaracter = "";
					List cellTempList2;
					try{
						cellTempList2 = (List) cellDataList.get(i+1);				
						rut1 = (HSSFCell) cellTempList2.get(j);
						cadenaCaracter = rut1.toString();				
					}catch(Exception e){}
					
					System.out.println("Afp   : " + cadenaCaracter);
					afp = "OK";
					matriz[1][2]=rut1.toString();		
				}
			}
			
			
			
			
			
			
			}
			}
			String hghg = "";
			}catch(Exception e){System.out.println("Error en Busqueda de patrones: " + e + " en archivo : " + fileName);}
			return matriz;
		}
	
	
		
		
		
		
		
		public String[][] DevuelveConductorFolio(String fileName, int hoja) {
			String matriz[][] = new String[2000][100];
			matriz[1][1] = "";
			matriz[1][2] = "";
			int hojasTotal = 0;

			List cellDataList = new ArrayList();



			try {

			FileInputStream fileInputStream = new FileInputStream(fileName);
			POIFSFileSystem fsFileSystem = new POIFSFileSystem(fileInputStream);
			HSSFWorkbook workBook = new HSSFWorkbook(fsFileSystem);
			FormulaEvaluator evaluator = workBook.getCreationHelper().createFormulaEvaluator();

			int i = hoja-1;
			hojasTotal = i;
			//for(int i=0;i<hojasTotal;i++){

			HSSFSheet hssfSheet = workBook.getSheetAt(i);
			Iterator rowIterator = hssfSheet.rowIterator();
			while (rowIterator.hasNext()){
			HSSFRow hssfRow = (HSSFRow) rowIterator.next();
			Iterator iterator = hssfRow.cellIterator();



			while (iterator.hasNext()){
			List cellTempList = new ArrayList();
			HSSFCell hssfCell = (HSSFCell) iterator.next();
			CellValue cellValue = evaluator.evaluate(hssfCell);
			String stringCellValue = hssfCell.toString();

			if(stringCellValue.trim().equals("")){}else{
			if(hssfCell.getCellType()==2){
			HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(workBook);
			CellValue cv = null;

			try {
			cv = fe.evaluate(hssfCell);

			if(cv.getCellType()==4){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getBooleanValue());
			}

			if(cv.getCellType()==2){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getStringValue());
			}

			if(cv.getCellType()==1){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getStringValue());
			}

			if(cv.getCellType()==0){
			hssfCell.setCellType(0);
			hssfCell.setCellValue(cv.getNumberValue());
			}

			} catch (Exception e) {matriz[1][1]= "";}
			}

			cellTempList.add(hssfCell);
			if(cellTempList.isEmpty()){}else{cellDataList.add(cellTempList);}

			}

			}



			}
			hssfSheet = null;
			rowIterator = null;

			//} // End For

			fileInputStream = null;
			fsFileSystem = null;
			workBook =null;
			evaluator = null;
			}catch (Exception e) {System.out.println("Error en 1 Ciclo Devuelve : " + fileName + "-" + e);}

			String ord ="",folio = "",total = "";








			try{

			int posibilidades = 0;
			cellDataList = cellDataList;
			HSSFCell rut1 = null;
			String matrizX[][] = new String[2000][100];

			int lineas = cellDataList.size(); // Cantidad de lineas
			for (int i = 0; i < cellDataList.size(); i++){
			List cellTempList = (List) cellDataList.get(i);

			for (int j = 0; j < cellTempList.size(); j++){
			HSSFCell hssfCell = (HSSFCell) cellTempList.get(j);
			String stringCellValue = hssfCell.toString();
			stringCellValue=stringCellValue.toUpperCase();


			

			posibilidades = stringCellValue.indexOf("FOLI");
			if (posibilidades== -1 ){matriz[1][1] = "";}else{
				List cellTempList2 = (List) cellDataList.get(i+1);
				rut1 = (HSSFCell) cellTempList2.get(j);
				matriz[1][1]=rut1.toString();
				System.out.println("Es : " + rut1.toString());
				j=999;
				i=999;

			}
			}
			}
			String hghg = "";
			}catch(Exception e){System.out.println("Error en Busqueda de patrones: " + e + " en archivo : " + fileName);}
			return matriz;
		}
	
	


		
		
		
		
		
		

		
		
		public String[][] DevuelveConductorFecha(String fileName, int hoja) {
			String matriz[][] = new String[2000][100];
			matriz[1][1] = "";
			matriz[1][2] = "";
			int hojasTotal = 0;

			List cellDataList = new ArrayList();



			try {

			FileInputStream fileInputStream = new FileInputStream(fileName);
			POIFSFileSystem fsFileSystem = new POIFSFileSystem(fileInputStream);
			HSSFWorkbook workBook = new HSSFWorkbook(fsFileSystem);
			FormulaEvaluator evaluator = workBook.getCreationHelper().createFormulaEvaluator();

			int i = hoja-1;
			hojasTotal = i;
			//for(int i=0;i<hojasTotal;i++){

			HSSFSheet hssfSheet = workBook.getSheetAt(i);
			Iterator rowIterator = hssfSheet.rowIterator();
			while (rowIterator.hasNext()){
			HSSFRow hssfRow = (HSSFRow) rowIterator.next();
			Iterator iterator = hssfRow.cellIterator();



			while (iterator.hasNext()){
			List cellTempList = new ArrayList();
			HSSFCell hssfCell = (HSSFCell) iterator.next();
			CellValue cellValue = evaluator.evaluate(hssfCell);
			String stringCellValue = hssfCell.toString();

			if(stringCellValue.trim().equals("")){}else{
			if(hssfCell.getCellType()==2){
			HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(workBook);
			CellValue cv = null;

			try {
			cv = fe.evaluate(hssfCell);

			if(cv.getCellType()==4){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getBooleanValue());
			}

			if(cv.getCellType()==2){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getStringValue());
			}

			if(cv.getCellType()==1){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getStringValue());
			}

			if(cv.getCellType()==0){
			hssfCell.setCellType(0);
			hssfCell.setCellValue(cv.getNumberValue());
			}

			} catch (Exception e) {matriz[1][1]= "";}
			}

			cellTempList.add(hssfCell);
			if(cellTempList.isEmpty()){}else{cellDataList.add(cellTempList);}

			}

			}



			}
			hssfSheet = null;
			rowIterator = null;

			//} // End For

			fileInputStream = null;
			fsFileSystem = null;
			workBook =null;
			evaluator = null;
			}catch (Exception e) {System.out.println("Error en 1 Ciclo Devuelve : " + fileName + "-" + e);}

			String ord ="",folio = "",total = "";








			try{

			int posibilidades = 0;
			cellDataList = cellDataList;
			HSSFCell rut1 = null;
			String matrizX[][] = new String[2000][100];

			int lineas = cellDataList.size(); // Cantidad de lineas
			for (int i = 0; i < cellDataList.size(); i++){
			List cellTempList = (List) cellDataList.get(i);

			for (int j = 0; j < cellTempList.size(); j++){
			HSSFCell hssfCell = (HSSFCell) cellTempList.get(j);
			String stringCellValue = hssfCell.toString();
			stringCellValue=stringCellValue.toUpperCase();


			
			
			
			
			
			
			
			
			
			
			
			

			posibilidades = stringCellValue.indexOf("FECHA");
			if (posibilidades== -1 ){}else{
				try{
					
					List cellTempList2 = (List) cellDataList.get(i+1);
					rut1 = (HSSFCell) cellTempList2.get(j);
					
					System.out.println("Archivo  : " + fileName + "  -   " +  rut1.toString());
					String dia = "";
					String mes = "";
					String ano = "";
					String cadena = rut1.toString().trim();   
					int largo = cadena.length();
					if(largo==15){
						//System.out.println("Se Procesara " + cadena);
						
						
						dia = stringCellValue.substring(10,12);
						mes = stringCellValue.substring(13,15);
						ano = "20" + stringCellValue.substring(16,18);
						//System.out.println("Dia " + dia);
						//System.out.println("Mes " + mes);
						//System.out.println("Ano " +ano);
						String fecha = dia+"/"+mes+"/"+ano;
						matriz[1][2]=fecha;
						matriz[1][1]="123";						
						j = 10000;
						i = 10000;
					}

					if (largo==11){
						dia = cadena.substring(0,2);
						mes = cadena.substring(3,6);
						ano = cadena.substring(7,11);
						mes=mes.toUpperCase();
						if(mes.trim().equals("ENE")){mes="01";}
						if(mes.trim().equals("FEB")){mes="02";}
						if(mes.trim().equals("MAR")){mes="03";}
						if(mes.trim().equals("ABR")){mes="04";}
						if(mes.trim().equals("MAY")){mes="05";}
						if(mes.trim().equals("JUN")){mes="06";}
						if(mes.trim().equals("JUL")){mes="07";}
						if(mes.trim().equals("AGO")){mes="08";}
						if(mes.trim().equals("SEP")){mes="09";}
						if(mes.trim().equals("OCT")){mes="10";}
						if(mes.trim().equals("NOV")){mes="11";}
						if(mes.trim().equals("DIC")){mes="12";}
						String fecha = dia+"/"+mes+"/"+ano;
						System.out.println("Es Valido : " + cadena + " MES : " + mes + " FIN :  " + fecha);
						matriz[1][2]=fecha;
						matriz[1][1]="123";
						j = 10000;
						i = 10000;
					}else{
						//System.out.println("NO Es Valido : " + cadena);
					}
					
					if (largo==10){
						dia = cadena.substring(0,2);
						mes = cadena.substring(3,5);
						ano = cadena.substring(6,10);
						mes=mes.toUpperCase();
						if(mes.trim().equals("ENE")){mes="01";}
						if(mes.trim().equals("FEB")){mes="02";}
						if(mes.trim().equals("MAR")){mes="03";}
						if(mes.trim().equals("ABR")){mes="04";}
						if(mes.trim().equals("MAY")){mes="05";}
						if(mes.trim().equals("JUN")){mes="06";}
						if(mes.trim().equals("JUL")){mes="07";}
						if(mes.trim().equals("AGO")){mes="08";}
						if(mes.trim().equals("SEP")){mes="09";}
						if(mes.trim().equals("OCT")){mes="10";}
						if(mes.trim().equals("NOV")){mes="11";}
						if(mes.trim().equals("DIC")){mes="12";}
						String fecha = dia+"/"+mes+"/"+ano;
						System.out.println("Es Valido : " + cadena + " MES : " + mes + " FIN :  " + fecha);
						matriz[1][2]=fecha;
						matriz[1][1]="123";
						j = 10000;
						i = 10000;
					}else{
						//System.out.println("NO Es Valido : " + cadena);
					}
					
				}catch(Exception e){}
			}
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			}
			}
			String hghg = "";
			}catch(Exception e){System.out.println("Error en Busqueda de patrones: " + e + " en archivo : " + fileName);}
			return matriz;
		}
	
	


		
		
		
		
		
		public String[][] DevuelveConductorTotales(String fileName, int hoja) {
			String matriz[][] = new String[2000][100];
			matriz[1][1] = "";
			matriz[1][2] = "";
			int hojasTotal = 0;

			List cellDataList = new ArrayList();



			try {

			FileInputStream fileInputStream = new FileInputStream(fileName);
			POIFSFileSystem fsFileSystem = new POIFSFileSystem(fileInputStream);
			HSSFWorkbook workBook = new HSSFWorkbook(fsFileSystem);
			FormulaEvaluator evaluator = workBook.getCreationHelper().createFormulaEvaluator();

			int i = hoja-1;
			hojasTotal = i;
			//for(int i=0;i<hojasTotal;i++){

			HSSFSheet hssfSheet = workBook.getSheetAt(i);
			Iterator rowIterator = hssfSheet.rowIterator();
			while (rowIterator.hasNext()){
			HSSFRow hssfRow = (HSSFRow) rowIterator.next();
			Iterator iterator = hssfRow.cellIterator();



			while (iterator.hasNext()){
			List cellTempList = new ArrayList();
			HSSFCell hssfCell = (HSSFCell) iterator.next();
			CellValue cellValue = evaluator.evaluate(hssfCell);
			String stringCellValue = hssfCell.toString();

			if(stringCellValue.trim().equals("")){}else{
			if(hssfCell.getCellType()==2){
			HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(workBook);
			CellValue cv = null;

			try {
			cv = fe.evaluate(hssfCell);

			if(cv.getCellType()==4){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getBooleanValue());
			}

			if(cv.getCellType()==2){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getStringValue());
			}

			if(cv.getCellType()==1){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getStringValue());
			}

			if(cv.getCellType()==0){
			hssfCell.setCellType(0);
			hssfCell.setCellValue(cv.getNumberValue());
			}

			} catch (Exception e) {matriz[1][1]= "";}
			}

			cellTempList.add(hssfCell);
			if(cellTempList.isEmpty()){}else{cellDataList.add(cellTempList);}

			}

			}



			}
			hssfSheet = null;
			rowIterator = null;

			//} // End For

			fileInputStream = null;
			fsFileSystem = null;
			workBook =null;
			evaluator = null;
			}catch (Exception e) {System.out.println("Error en 1 Ciclo Devuelve : " + fileName + "-" + e);}

			String ord ="",folio = "",total = "";








			try{

			int posibilidades = 0;
			cellDataList = cellDataList;
			HSSFCell rut1 = null;
			String matrizX[][] = new String[2000][100];

			int lineas = cellDataList.size(); // Cantidad de lineas
			for (int i = 0; i < cellDataList.size(); i++){
			List cellTempList = (List) cellDataList.get(i);

			for (int j = 0; j < cellTempList.size(); j++){
			HSSFCell hssfCell = (HSSFCell) cellTempList.get(j);
			String stringCellValue = hssfCell.toString();
			stringCellValue=stringCellValue.toUpperCase();


			

			// TOTAL
           
			HSSFCell totalDebeFind = null;
			HSSFCell totalHaberFind = null;
			
            int posibilidades5 = stringCellValue.indexOf("TOTAL");
            if (posibilidades5== -1 ){}else{
            	
            	
            		List cellTempList2 = (List) cellDataList.get(i+1);
            		totalDebeFind = (HSSFCell) cellTempList2.get(j);
            
            		cellTempList2 = (List) cellDataList.get(i+2);
            		totalHaberFind = (HSSFCell) cellTempList2.get(j);
            	
                   // totalDebeFind=(HSSFCell) cellTempList.get(j+1);         
                   // totalHaberFind=(HSSFCell) cellTempList.get(j+2);
                    
                    
                    
                    if (totalDebeFind.getCellType()==0){
                            double hh = totalDebeFind.getNumericCellValue();
                            long g = Math.round(hh);
                            
                            matriz[1][2]=String.valueOf(g);
                            matriz[1][1]="123";
                            
                            
                            j = 10000;
                            i = 10000;
                    }else{
                            matriz[1][2]=totalDebeFind.toString();
                            
                            matriz[1][2] = matriz[1][2].replace(".", "");
                            matriz[1][2] = matriz[1][2].replace(",", "");
                            matriz[1][2] = matriz[1][2].replace("$", "");
                            matriz[1][2] = matriz[1][2].replace(" ", "");
                            matriz[1][2] = matriz[1][2].replace("-", "");
                            
                            
                            matriz[1][1]="123";
                            
                            j = 10000;
                            i = 10000;
                    }
                    
                    if (totalHaberFind.getCellType()==0){
                            double hh = totalHaberFind.getNumericCellValue();
                            long g = Math.round(hh);
                            
                            matriz[1][3]=String.valueOf(g);
                            matriz[1][1]="123";
                            
                            
                            j = 10000;
                            i = 10000;
                    }else{
                            matriz[1][3]=totalHaberFind.toString();
                            matriz[1][3] = matriz[1][3].replace(".", "");
                            matriz[1][3] = matriz[1][3].replace(",", "");
                            matriz[1][3] = matriz[1][3].replace("$", "");
                            matriz[1][3] = matriz[1][3].replace(" ", "");
                            matriz[1][3] = matriz[1][3].replace("-", "");
                            matriz[1][1]="123";
                            
                            j = 10000;
                            i = 10000;
                    }
                    
                    
            }

			
			
			
			
			
			
			
			
			
			}
			}
			String hghg = "";
			}catch(Exception e){System.out.println("Error en Busqueda de patrones: " + e + " en archivo : " + fileName);}
			return matriz;
		}
		
		
		
		
		public String[][] DevuelveConductorRutnomb(String fileName, int hoja) {
			String matriz[][] = new String[2000][100];
			matriz[1][1] = "";
			matriz[1][2] = "";
			int hojasTotal = 0;

			List cellDataList = new ArrayList();



			try {

			FileInputStream fileInputStream = new FileInputStream(fileName);
			POIFSFileSystem fsFileSystem = new POIFSFileSystem(fileInputStream);
			HSSFWorkbook workBook = new HSSFWorkbook(fsFileSystem);
			FormulaEvaluator evaluator = workBook.getCreationHelper().createFormulaEvaluator();

			int i = hoja-1;
			hojasTotal = i;
			//for(int i=0;i<hojasTotal;i++){

			HSSFSheet hssfSheet = workBook.getSheetAt(i);
			Iterator rowIterator = hssfSheet.rowIterator();
			while (rowIterator.hasNext()){
			HSSFRow hssfRow = (HSSFRow) rowIterator.next();
			Iterator iterator = hssfRow.cellIterator();



			while (iterator.hasNext()){
			List cellTempList = new ArrayList();
			HSSFCell hssfCell = (HSSFCell) iterator.next();
			CellValue cellValue = evaluator.evaluate(hssfCell);
			String stringCellValue = hssfCell.toString();

			if(stringCellValue.trim().equals("")){}else{
			if(hssfCell.getCellType()==2){
			HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(workBook);
			CellValue cv = null;

			try {
			cv = fe.evaluate(hssfCell);

			if(cv.getCellType()==4){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getBooleanValue());
			}

			if(cv.getCellType()==2){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getStringValue());
			}

			if(cv.getCellType()==1){
			hssfCell.setCellType(1);
			hssfCell.setCellValue(cv.getStringValue());
			}

			if(cv.getCellType()==0){
			hssfCell.setCellType(0);
			hssfCell.setCellValue(cv.getNumberValue());
			}

			} catch (Exception e) {matriz[1][1]= "";}
			}

			cellTempList.add(hssfCell);
			if(cellTempList.isEmpty()){}else{cellDataList.add(cellTempList);}

			}

			}



			}
			hssfSheet = null;
			rowIterator = null;

			//} // End For

			fileInputStream = null;
			fsFileSystem = null;
			workBook =null;
			evaluator = null;
			}catch (Exception e) {System.out.println("Error en 1 Ciclo Devuelve : " + fileName + "-" + e);}

			String ord ="",folio = "",total = "";








			try{
String a = "NO";
String b = "NO";
			int posibilidades = 0;
			cellDataList = cellDataList;
			HSSFCell rut1 = null;
			String matrizX[][] = new String[2000][100];

			int lineas = cellDataList.size(); // Cantidad de lineas
			for (int i = 0; i < cellDataList.size(); i++){
			List cellTempList = (List) cellDataList.get(i);

			for (int j = 0; j < cellTempList.size(); j++){
			HSSFCell hssfCell = (HSSFCell) cellTempList.get(j);
			String stringCellValue = hssfCell.toString();
			stringCellValue=stringCellValue.toUpperCase();


			

			

             int posibilidades4 = stringCellValue.indexOf("RUT");
             if (posibilidades4== -1 ){}else{
                             try{
							 List cellTempList2 = (List) cellDataList.get(i+1);
							rut1 = (HSSFCell) cellTempList2.get(j);
                         		
                             
                             matriz[1][2]=rut1.toString();
                             matriz[1][1]="123";
                             
                             
                             matriz[1][2] = matriz[1][2].replace("-", "");
                             matriz[1][2] = matriz[1][2].replace(".", "");
                             matriz[1][2] = matriz[1][2].replace(";", "");
                             matriz[1][2] = matriz[1][2].replace(",", "");
                             
                             matriz[1][2] = matriz[1][2].trim().replace("E7","");
                             matriz[1][2] = matriz[1][2].trim().replace("E8","");
                             matriz[1][2] = matriz[1][2].trim().replace(" ","");
                             int largo = matriz[1][2].trim().length();
                             
                             String digito = matriz[1][2].trim().substring(largo-1,largo);                                                          
                             String rut  = matriz[1][2].trim().substring(0 , largo-1);
                             
                             matriz[1][2] = rut.trim();
                             matriz[1][3] = digito.trim();                             
                             b = "SI";
                            
                             
                             }catch(Exception e){System.out.println("Error Obtencion rut y Div : " + e);}
             }               
             
             //posibilidades4 = stringCellValue.indexOf("A NOMBRE D");
             posibilidades4 = stringCellValue.indexOf("IMPONENT");
             if (posibilidades4== -1 ){}else{
                             try{
                            	 List cellTempList2 = (List) cellDataList.get(i+1);
     							rut1 = (HSSFCell) cellTempList2.get(j);
                             //rut1 = (HSSFCell) cellTempList.get(j+1);
                             matriz[1][4]=rut1.toString();
                             matriz[1][1]="1";
                                                     
                             matriz[1][2]=matriz[1][2].replace("'", "");
                             a = "SI";
                            
                             
                             }catch(Exception e){System.out.println("Error Obtencion iMPONENTE : " + e);}
             }

			
			if(a.trim().equals("SI") && b.trim().equals("SI") ){			
			i = 99999;
			j = 9999;
			}
			}
			}
			String hghg = "";
			}catch(Exception e){System.out.println("Error en Busqueda de patrones: " + e + " en archivo : " + fileName);}
			return matriz;
		}
	
	

		
}