package com.mkyong.common.action;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.opensymphony.xwork2.ActionSupport;

public class FileUploadAction extends ActionSupport{

	private File fileUpload;
	private String fileUploadContentType;
	private String fileUploadFileName;
	private String readOut = "";

	public String getFileUploadContentType() {
		return fileUploadContentType;
	}

	public void setFileUploadContentType(String fileUploadContentType) {
		this.fileUploadContentType = fileUploadContentType;
	}

	public String getFileUploadFileName() {
		return fileUploadFileName;
	}

	public void setFileUploadFileName(String fileUploadFileName) {
		this.fileUploadFileName = fileUploadFileName;
	}

	public File getFileUpload() {
		return fileUpload;
	}

	public void setFileUpload(File fileUpload) {
		this.fileUpload = fileUpload;
	}
	

	public String getReadOut() {
		return readOut;
	}

	public void setReadOut(String readOut) {
		this.readOut = readOut;
	}

	public String execute() throws Exception{
		FileInputStream file = new FileInputStream(fileUpload);
		Workbook workbook = new XSSFWorkbook(file);
		
		Sheet sheet = workbook.getSheetAt(0);

		Map<Integer, List<String>> data = new HashMap<Integer, List<String>>();
		int i = 0;
		for (Row row : sheet) {
		    data.put(i, new ArrayList<String>());
		    for (Cell cell : row) {
		        switch (cell.getCellType()) {
		            case STRING: readOut = readOut + cell.getStringCellValue(); break;
		            case NUMERIC: readOut = readOut + String.valueOf(Double.valueOf(cell.getNumericCellValue()).intValue()); break;
//		            case BOOLEAN: ... break;
//		            case FORMULA: ... break;
		            default: data.get(Integer.valueOf(i)).add(" ");
		        }
		    }
		    i++;
		}
		System.out.println(readOut);
		workbook.close();
		
		return SUCCESS;
		
	}
	
	public String display() {
		return NONE;
	}
	
}