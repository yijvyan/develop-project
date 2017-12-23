package com.yijvyan.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.output.Format;
import org.jdom2.output.XMLOutputter;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
/**
 * 
 * @author yijvyan 2017.12.23
 * thanks to chenyunlin1342 from CSDN
 * http://blog.csdn.net/chenyunlin1342/article/details/49737721
 *
 */
public class EToXDemo {
	
	public EToXDemo() {}

	public static void main(String[] args) throws IOException {
		EToXDemo demo = new EToXDemo();
		demo.change();
		System.out.println("Mission Complete!");
	}

	private void change() throws IOException {
		Workbook wb = null;
		Element root = new Element("sheet");
		
		Document doc = new Document(root);
		InputStream in = new FileInputStream("e:\\excel.xls");
		
		try {
			wb = Workbook.getWorkbook(in);
			Sheet sheet = wb.getSheet(0);
				
			int columns = sheet.getColumns();
			int rows = sheet.getRows();
			
			for (int i = 0; i < rows; i++) {
				Element elements = new Element("tr");
				for (int j = 0; j < columns; j++) {
					Cell cell = sheet.getCell(j, i);
					
					elements.addContent(new Element("cell").setText(cell.getContents()));
					root.addContent(elements.detach());
				}
			}
			
			Format format = Format.getPrettyFormat();
			XMLOutputter xmlOut = new XMLOutputter(format);
			xmlOut.output(doc, new FileOutputStream("e:\\eToXml.xml"));
		} catch (BiffException e) {
			e.printStackTrace();
		}
		
	}
}
