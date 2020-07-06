package main;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;


public class MainApplication {
	public static int majorCredit;
	public static int majorPassCredit;
	public static double majorGrade;
	
	public static int subMajorCredit;
	public static int subMajorPassCredit;
	public static double subMajorGrade;
	
	public static int normalCredit;
	public static int normalPassCredit;
	public static double normalGrade;
	
	public static int courseIdx;
	public static int creditIdx;
	public static int gradeIdx;
	
	public static void main(String[] args) {
		
		BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
		System.out.print("파일 경로와 이름을 입력해주세요 > ");
		
		try {
			FileInputStream fr = new FileInputStream(br.readLine());
			HSSFWorkbook workbook = new HSSFWorkbook(fr);
			HSSFSheet sheet = workbook.getSheetAt(0);
			int rows = sheet.getPhysicalNumberOfRows();
			
			for (int rowIdx = 0; rowIdx < 1; rowIdx++) {
				
				HSSFRow row = sheet.getRow(rowIdx);
				int cells = row.getPhysicalNumberOfCells();
				for (int colIdx = 0; colIdx < cells; colIdx++) {
					if(row.getCell(colIdx).toString().equals("이수구분")) {
						courseIdx = colIdx;
					}
					if(row.getCell(colIdx).toString().equals("학점")) {
						creditIdx = colIdx;
					}
					if(row.getCell(colIdx).toString().equals("등급")) {
						gradeIdx = colIdx;
					}
				}
			}

			for (int rowIdx = 1; rowIdx < rows; rowIdx++) {
				
				HSSFRow row = sheet.getRow(rowIdx);
				String course = row.getCell(courseIdx).toString();
				int credit = Integer.parseInt(String.valueOf(row.getCell(creditIdx)));
				String grade = String.valueOf(row.getCell(gradeIdx));
				
				if(grade.equals("F"))
					continue;
				
				if(course.equals("전필") || course.equals("전선")) {
					if(grade.equals("P")) {
						majorPassCredit += credit;
					} else {
						majorCredit += credit;
						majorGrade += (gradeCalc(grade) * credit);
					}
				} else if(course.equals("부전")) {
					if(grade.equals("P")) {
						subMajorPassCredit += credit;
					} else {
						subMajorCredit += credit;
						subMajorGrade += (gradeCalc(grade) * credit);
					}
				} else {
					if(grade.equals("P")) {
						normalPassCredit += credit;
					} else {
						normalCredit += credit;
						normalGrade += (gradeCalc(grade) * credit);
					}
				}
			}
			
			System.out.println(majorCredit + " " + majorGrade + " " + majorPassCredit);
			System.out.println(subMajorCredit + " " + subMajorGrade + " " + subMajorPassCredit);
			System.out.println(normalCredit + " " + normalGrade + " " + normalPassCredit);
			
			workbook.close();
			
		} catch (IOException e) {
			System.out.println(e);
		}
	}
	
	public static double gradeCalc(String grade) {
		
		if(grade.equals("A+")) {
			return 4.5;
		}
		if(grade.equals("A0")) {
			return 4.0;
		}
		if(grade.equals("B+")) {
			return 3.5;
		}
		if(grade.equals("B0")) {
			return 3.0;
		}
		if(grade.equals("C+")) {
			return 2.5;
		}
		if(grade.equals("C0")) {
			return 2.0;
		}
		if(grade.equals("D+")) {
			return 1.5;
		}
		
		return 1.0;
	}
}
