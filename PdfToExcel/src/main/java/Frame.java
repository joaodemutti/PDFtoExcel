package main.java;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.FileDialog;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.encryption.AccessPermission;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;



public class Frame extends JFrame implements ActionListener {
	
		JButton button = new JButton("Choose a file");
		
		
	public Frame() {
		setDefaultCloseOperation(EXIT_ON_CLOSE);
		setSize(230, 120);
		setLayout(null);
		setLocationRelativeTo(null);
	
		button.setBounds(50,30,100,20);
		button.addActionListener(this);
		button.setFocusable(false);
		button.setHorizontalAlignment(JButton.CENTER);
		button.setVerticalAlignment(JButton.CENTER);
		button.setBounds(35,45,140,20);
		add(button);
		
		var label = new JLabel("Simple PDF to Excel.");
		label.setBounds(10,10, 200, 20);
		add(label);
		setVisible(true);
	}

	public void ExportExcel(String[] pages) {
		JOptionPane.showMessageDialog(this, "You will be asked to save your file", "", JOptionPane.PLAIN_MESSAGE);
		FileDialog fd = new FileDialog(this,"Save your Excel file",FileDialog.SAVE);
		fd.setFile("*.xls");
		fd.setVisible(true);
		
		File file = fd.getFiles()[0];
		if(file == null)
			JOptionPane.showMessageDialog(this, "You didn't saved your file", "", JOptionPane.PLAIN_MESSAGE);
		else {
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet("FirstSheet");  

			var i = 0;
			for(var page : pages) {
				HSSFRow rowhead = sheet.createRow((short)i);
	            var cell = rowhead.createCell(0);
	            cell.setCellValue(page);
	            i++;
			}
            try {
            	FileOutputStream fileOut = new FileOutputStream(file);
                workbook.write(fileOut);
                fileOut.flush();
                fileOut.close();
                workbook.close();
                JOptionPane.showMessageDialog(this, "Your file was saved successfully", "", JOptionPane.PLAIN_MESSAGE);
            }catch(IOException ex) {
    			JOptionPane.showMessageDialog(this, ex, "Error saving excel file", JOptionPane.PLAIN_MESSAGE);
    		}
            
		}
	}
	
	public void ExtractPDF(File file) throws IOException{
		
		PDDocument document = Loader.loadPDF(file);
		
		AccessPermission ap = document.getCurrentAccessPermission();
        if (!ap.canExtractContent())
        {
            throw new IOException("You do not have permission to extract text");
        }
        
        PDFTextStripper stripper = new PDFTextStripper();

        stripper.setSortByPosition(true);

        String[] pages = new String[ document.getNumberOfPages()];
        
        for (int p = 0; p < document.getNumberOfPages(); p++)
        {
        	stripper.setStartPage(p+1);
            stripper.setEndPage(p+1);

            String text = stripper.getText(document);
            pages[p] = text;
           
        }
        
        ExportExcel(pages);
        
	}
	
	public void ChooseFile() {
		FileDialog fd = new FileDialog(this,"Choose a pdf file",FileDialog.LOAD);
		fd.setFile("*.pdf");
		fd.setVisible(true);
		var files = fd.getFiles();
		try {
		for(var file : files) {
			ExtractPDF(file);
		}
		}catch(IOException ex) {
			JOptionPane.showMessageDialog(this, ex, "Error loading pdf file", JOptionPane.PLAIN_MESSAGE);
		}
	}
	
	@Override
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		if(e.getSource()==button) {
			ChooseFile();
		}
	}
}
