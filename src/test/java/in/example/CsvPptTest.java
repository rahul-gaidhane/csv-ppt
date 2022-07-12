package in.example;

import java.awt.Rectangle;
import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.junit.jupiter.api.Test;

public class CsvPptTest {
	
	private static final String COMMA_DELIMITER = ",";

	@Test
	public void test() throws FileNotFoundException, IOException {
		List<List<String>> records = new ArrayList<>();
		
		try (BufferedReader br = new BufferedReader(new FileReader("src/main/resources/source.csv"))) {
		    String line;
		    while ((line = br.readLine()) != null) {
		        String[] values = line.split(COMMA_DELIMITER);
		        records.add(Arrays.asList(values));
		    }
		}
		
		records.forEach(rec -> System.out.println("Rec : " + rec));
		
		XMLSlideShow slide = new XMLSlideShow();
		XSLFSlide slide1 =  slide.createSlide();
		
		XSLFTable tbl = slide1.createTable();
		tbl.setAnchor(new Rectangle(50, 50, 450, 300));
		XSLFTableRow headerRow = tbl.addRow();
		headerRow.setHeight(25);
		
		XSLFTableCell headerCell1 = headerRow.addCell();
	    XSLFTextParagraph headerPara1 = headerCell1.addNewTextParagraph();
	    headerPara1.setTextAlign(TextParagraph.TextAlign.CENTER);
	    XSLFTextRun headerText1 = headerPara1.addNewTextRun();
	    headerText1.setText("Name");
	    
	    XSLFTableCell headerCell2 = headerRow.addCell();
	    XSLFTextParagraph headerPara2 = headerCell2.addNewTextParagraph();
	    headerPara2.setTextAlign(TextParagraph.TextAlign.CENTER);
	    XSLFTextRun headerText2 = headerPara2.addNewTextRun();
	    headerText2.setText("Grade");
	    
	    XSLFTableCell headerCell3 = headerRow.addCell();
	    XSLFTextParagraph headerPara3 = headerCell3.addNewTextParagraph();
	    headerPara3.setTextAlign(TextParagraph.TextAlign.CENTER);
	    XSLFTextRun headerText3 = headerPara3.addNewTextRun();
	    headerText3.setText("Marks");
	    
		records.forEach(rec -> {
			XSLFTableRow newRow = tbl.addRow();
			for(int i = 0; i < 3; i++) {
				XSLFTableCell th = newRow.addCell();
			    XSLFTextParagraph p = th.addNewTextParagraph();
			    p.setTextAlign(TextParagraph.TextAlign.CENTER);
			    XSLFTextRun r = p.addNewTextRun();
			    r.setText(rec.get(i));
			}
		});
		
		
		
		FileOutputStream out = new FileOutputStream("src/main/resources/slide.pptx");
		slide.write(out);
		out.close();
		slide.close();
		
	}
}
