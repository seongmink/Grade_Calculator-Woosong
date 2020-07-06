package main;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileWriter;
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

			init(sheet);
			calc(sheet);
			printResult();
			
			br.close();
			workbook.close();
		} catch (IOException e) {
			System.out.println(e);
		}
	}
	
	public static void init(HSSFSheet sheet) {
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
	}
	
	public static void calc(HSSFSheet sheet) {
		int rows = sheet.getPhysicalNumberOfRows();
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
	}
	
	public static void printResult() {
		
		StringBuilder sb = new StringBuilder();
		
		int totalMajorCredit = majorCredit + majorPassCredit;
		double totalMajorGrade = Math.round((majorGrade / majorCredit) * 100) / 100.0;
		int totalSubMajorCredit = subMajorCredit + subMajorPassCredit;
		double totalSubMajorGrade = Math.round((subMajorGrade / subMajorCredit) * 100) / 100.0;
		int totalNormalCredit = normalCredit + normalPassCredit;
		double totalNormalGrade = Math.round((normalGrade / normalCredit) * 100) / 100.0;
		int totalCredit = majorCredit + subMajorCredit + normalCredit;
		int totalPassCredit = majorPassCredit + subMajorPassCredit + normalPassCredit;
		double totalGrade = majorGrade + subMajorGrade + normalGrade;
		
		sb.append("계산한 결과는 다음과 같습니다.(소수점 세번째 자리에서 반올림)\n\n");
		sb.append("전공 취득 학점 : ").append(totalMajorCredit + "학점\n");
		sb.append("전공 평균 학점 : ").append(totalMajorGrade + "\n");
		sb.append("부전 취득 학점 : ").append(totalSubMajorCredit + "학점\n");
		sb.append("부전 평균 학점 : ").append(totalSubMajorGrade + "\n");
		sb.append("교양 및 일선 취득 학점 : ").append(totalNormalCredit + "학점\n");
		sb.append("교양 및 일선 평균 학점 : ").append(totalNormalGrade + "\n");
		sb.append("전공을 제외한 취득 학점 : ").append(totalSubMajorCredit + totalNormalCredit + "\n");
		sb.append("전공을 제외한 평균 학점 : ").append(Math.round(totalGrade - majorGrade) / (totalCredit - majorCredit) * 100 / 100.0 +"\n");
		sb.append("총 취득 학점 : ").append(totalCredit + totalPassCredit + "학점\n");
		sb.append("총 평균 학점 : ").append(Math.round((totalGrade / totalCredit) * 100) / 100.0 + "\n");
		
		System.out.println(sb.toString());
		
		makeResult(sb);
	}
	
	public static void makeResult(StringBuilder sb) {
		
		try {
			String filePath = "C:\\Users\\KSM\\Desktop\\";
			String fileName = "내 학점 결과.txt";
			BufferedWriter bw = new BufferedWriter(new FileWriter(filePath + fileName, true));
			
			bw.write(sb.toString());
			bw.close();
			System.out.println("결과 파일이 정상적으로 생성되었습니다.");
			System.out.println("위치 > " + filePath + fileName);
		} catch (Exception e) {
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
