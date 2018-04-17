package model;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class RaterChecker {

	public static void main(String[] args) throws IOException {
	
		
		
		FileInputStream fis = new FileInputStream(new File("C://Data//RFATEMP.xls"));
		
		
		HSSFWorkbook wb = null;
		
		try {
			
			wb = new HSSFWorkbook(fis);
			
			System.out.println("Calculations Started...");
			
			long startTime = System.currentTimeMillis();
			
			HSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
			
			long endTime = System.currentTimeMillis();
			
			long duration = (endTime - startTime);
			
			System.out.println("Calculations done in " + duration + " milliseconds.");
				
		} catch (Exception e) {
			
			e.printStackTrace();
			
		} 
		
		
		/*long startTime = System.currentTimeMillis();
		
		Workbook wb = new XSSFWorkbook(fis);
		FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
		evaluator.evaluateAll();
		
		
		
		long endTime = System.currentTimeMillis();
		
		long duration = (endTime - startTime);
		
		System.out.println(duration + " milliseconds!!!");*/
		
		
		

	}

}
