import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Pattern;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.filechooser.FileNameExtensionFilter;

import jxl.Cell;
import jxl.CellView;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExpList {
	
	private ArrayList<String> labels = new ArrayList<String>();
	private Map<String, ArrayList<String[]>> sheets = new HashMap<String, ArrayList<String[]>>();
	
	private File filter = new File("filter.csv");
	private File columns = new File("spalten.csv");

	public static void main(String[] args) {
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
				| UnsupportedLookAndFeelException e) {
			e.printStackTrace();
		}
		new ExpList();
	}
	
	public ExpList(){
		
		if(!filter.exists() || !columns.exists()){
			JOptionPane.showMessageDialog(null, "filter.csv und spalten.csv müssen existieren!");
			System.exit(0);
		}
		
		JFileChooser fileOpener = new JFileChooser(System.getProperty("user.home") + "\\Desktop");
		fileOpener.addChoosableFileFilter(new FileNameExtensionFilter("txt", "csv", "xls"));
		fileOpener.setFileSelectionMode(JFileChooser.FILES_ONLY);
		fileOpener.setDialogTitle("Wähle die aus SAP exportiere Datei");
		fileOpener.setApproveButtonText("Bearbeiten");
		
		int returnValue = fileOpener.showOpenDialog(null);
		
		if(returnValue == JFileChooser.APPROVE_OPTION){
			try{
				System.out.println("Read " + fileOpener.getSelectedFile().getAbsolutePath());
				SapExport expList = new SapExport(fileOpener.getSelectedFile());
				expList.filterLines(Filter.importFilters(filter));
				expList.filterColumns(readLines(columns));
				System.out.println("After applying filters: " + expList.data.size());
				expList.removeDuplicates();
				System.out.println("After removing duplicates: " + expList.data.size());
				
				File exclude = new File("exclude");
				if(exclude.exists() && exclude.isDirectory()){
					File[] excludes = exclude.listFiles();
					for(int i=0;i<excludes.length;i++){
						System.out.println("Excluding " + excludes[i].getName());
						expList.filterSapExport(new SapExport(excludes[i]), "Material");
					}
				}
				
				System.out.println("Without exludes folder: " + expList.data.size());
				System.out.println("Sorting list alphabetically..");
				expList.sort();
				
				/*int index = JOptionPane.showOptionDialog(
						null,
						"Nach welcher Spalte sollen die Daten separiert werden?",
						"Daten Separation",
						JOptionPane.YES_OPTION,
						JOptionPane.QUESTION_MESSAGE,
						null,
						expList.labels.toArray(new String[expList.labels.size()]),
						expList.labels.size() >= 3 ? expList.labels.get(2) : "" //default "LOrt"
				);
				
				if(index == -1){
					JOptionPane.showMessageDialog(null, "Wähle eine Spalte aus!");
					System.exit(0);
				}*/
				labels = expList.labels;
				
				int lortIndex = -1;
				for(int j=0;j<labels.size();j++){
					if(labels.get(j).equalsIgnoreCase("LOrt")){
						lortIndex = j;
					}
				}
				
				sheets = expList.seperateIntoSheetMap(sheets, expList.labels.get(lortIndex));
			}catch(IOException e){
				e.printStackTrace();
				JOptionPane.showMessageDialog(null, "Die Datei konnte nicht gelesen werden! Bitte versuche es erneut.");
				System.exit(0);
			}
			
			try {
				JFileChooser fileSaver = new JFileChooser("U:\\KSAapotheke\\Dienstleistung\\OffizinPha\\Verfall");
				fileSaver.addChoosableFileFilter(new FileNameExtensionFilter("xls", "txt"));
				fileSaver.setFileSelectionMode(JFileChooser.FILES_ONLY);
				fileSaver.setDialogTitle("Wo soll die Datei gespeichert werden?");
				fileSaver.setApproveButtonText("Speichern");
				fileSaver.setSelectedFile(new File("verfall.xls"));
				returnValue = fileSaver.showOpenDialog(null);
				
				if(returnValue != JFileChooser.APPROVE_OPTION){
					System.exit(0);
				}
				
				WritableWorkbook workbook = Workbook.createWorkbook(fileSaver.getSelectedFile());
				int c = 0;
				//System.out.println(sheets.keySet().size());
				
				WritableFont headerFont = new WritableFont(WritableFont.createFont("Arial"), WritableFont.DEFAULT_POINT_SIZE, WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE, Colour.WHITE);
				WritableFont normalFont = new WritableFont(WritableFont.createFont("Arial"), WritableFont.DEFAULT_POINT_SIZE, WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE, Colour.BLACK);
				
				Colour gray = Colour.GRAY_80;
				Colour light_gray = Colour.GRAY_25;
				 
				for (String key : sheets.keySet()) {
					
					System.out.println("Add sheet " + key);
					WritableSheet sheet = workbook.createSheet(key, c);
					
					//data
					for(int i=0;i<sheets.get(key).size();i++){
						for(int j=0;j<sheets.get(key).get(i).length;j++){
							Label label = new Label(j, i+1/*header line*/, sheets.get(key).get(i)[j]);
							
							WritableCellFormat format = new WritableCellFormat();
							format.setWrap(true);
							format.setBackground(i%2 == 0 ? Colour.WHITE : light_gray);
							format.setFont(normalFont);
							label.setCellFormat(format);
							sheet.addCell(label);
						}
					}
					//header after, autosize
					for(int k=0;k<labels.size();k++){
						Label label = new Label(k, 0, labels.get(k)); 
						
						WritableCellFormat format = new WritableCellFormat();
						format.setWrap(true);
						format.setBackground(gray);
						format.setFont(headerFont);
						label.setCellFormat(format);
						sheet.addCell(label);
						
//						CellView cell = sheet.getColumnView(k);
//						cell.setAutosize(true);
//						sheet.setColumnView(k, cell);
						
//						System.out.println("set " + k + " 0 to " + labels.get(k));
					}
					
					sheetAutoFitColumns(sheet);
					c++;
				}
				
				workbook.write(); 
				workbook.close();
				
				JOptionPane.showMessageDialog(null, "Die Datei wurde erfolgreich gespeichert!", "Yeah!", JOptionPane.INFORMATION_MESSAGE, null);
				
			} catch (IOException e) {
				JOptionPane.showMessageDialog(null, "Die Datei ist möglicherweise bereits geöffnet! (z.B. in Excel)");
				e.printStackTrace();
			} catch (RowsExceededException e) {
				JOptionPane.showMessageDialog(null, "Die Datei hat zuviele Zeilen!");
				e.printStackTrace();
			} catch (WriteException e) {
				JOptionPane.showMessageDialog(null, "Die Datei kann nicht beschrieben werden!");
				e.printStackTrace();
			}
			
		}else{
			System.exit(0);
		}
	}
	
	private void sheetAutoFitColumns(WritableSheet sheet) {
	    for (int i = 0; i < sheet.getColumns(); i++) {
	        Cell[] cells = sheet.getColumn(i);
	        int longestStrLen = -1;

	        if (cells.length == 0)
	            continue;

	        /* Find the widest cell in the column. */
	        for (int j = 0; j < cells.length; j++) {
	            if ( cells[j].getContents().length() > longestStrLen ) {
	                String str = cells[j].getContents();
	                if (str == null || str.isEmpty())
	                    continue;
	                longestStrLen = str.trim().length();
	            }
	        }

	        /* If not found, skip the column. */
	        if (longestStrLen == -1) 
	            continue;

	        CellView cv = sheet.getColumnView(i);
	        cv.setSize(longestStrLen * 280 + 150); /* Every character is 256 units wide, so scale it. */
	        sheet.setColumnView(i, cv);
	    }
	}
	
	private String[] readLines(File f) throws IOException{
		ArrayList<String> lines = new ArrayList<String>();
		
		FileReader fr = new FileReader(f);
		BufferedReader br = new BufferedReader(fr);
		String currentLine;
		
		while ((currentLine = br.readLine()) != null) {
			if(currentLine.trim().equals("")){continue;}
			
			lines.add(currentLine.trim());
		}
		
		br.close();
		
		return lines.toArray(new String[lines.size()]);
	}
	
	static class Filter{
		String column;
		String regex;
		
		Filter(String column, String regex){
			this.column = column;
			this.regex = regex;
		}
		
		public static Filter[] importFilters(File f) throws IOException{
			ArrayList<Filter> filters = new ArrayList<Filter>();
			
			FileReader fr = new FileReader(f);
			BufferedReader br = new BufferedReader(fr);
			String currentLine;
			int line = 0;
			
			while ((currentLine = br.readLine()) != null) {
				if(currentLine.trim().equals("")){continue;}
				
				String[] cols = currentLine.trim().split(";");
				for(int i=0;i<cols.length;i++){
					cols[i] = cols[i].trim();
				}
				
				if(line != 0){
					filters.add(new Filter(cols[0], cols[1]));
				}
				
				line++;
			}
			
			br.close();
			
			return filters.toArray(new Filter[filters.size()]);
		}
	}
	
	
	class SapExport{
		public ArrayList<String> labels = new ArrayList<String>();
		public ArrayList<String[]> data = new ArrayList<String[]>();
		
		SapExport(ArrayList<String> labels, ArrayList<String[]> data){
			this.labels = labels;
			this.data = data;
		}
		
		SapExport(File f) throws IOException{
			FileReader fr = new FileReader(f);
			BufferedReader br = new BufferedReader(fr);
			String currentLine;
			int line = 0; //The read line numbers
			int actualLine = -1; //The actual file line number
			boolean started = false;
			
			while ((currentLine = br.readLine()) != null) {
				actualLine++; //ensure this is incremented every time
				
				currentLine = currentLine.trim().replaceAll("[^A-z0-9 ./%&*\t]", "");
				
				
				if(!started && currentLine.contains("Material")){
					System.out.println("Found \"Material\" on line #" + actualLine + ", starting parsing from here on");
					started = true;
				}
				
				if(currentLine.equals("") || !started){
					System.out.println("Skipped line #" + actualLine);
					continue;
				}
				
				String[] cols = currentLine.split("	");
				for(int i=0;i<cols.length;i++){
					if(line == 0){
						System.out.println("Adding label " + cols[i].trim());
					}
					if(line != 0 && labels.get(i).equalsIgnoreCase("LOrt") && cols[i].trim().length() == 0){
						cols[i] = "Ohne Lagerort";
					}else{
						cols[i] = cols[i].trim();
					}
				}
				
				if(line == 0){
					labels.addAll(Arrays.asList(cols));
				}else if(labels.size() == cols.length){
					data.add(cols);
				}
				
				line++;
			}
			
			if(!started){
				JOptionPane.showMessageDialog(null, "Die ausgewählte Datei enthält kein \"Material\" und kann daher nicht eingelesen werden!\nMöglicherweise wurde die Datei nicht korrekt exportiert?");
				System.exit(0);
			}
			
			br.close();
		}
		
		public void filterLines(Filter[] filters){
			//reverse order => not affecting indices
			System.out.println(data.size() + " " + (data.size()-1));
			for(int i= (data.size()-1);i>-1;i--){
				//System.out.println(i + "/" + data.size() +": "+data.get(i)[0]);
				String[] col = data.get(i);
				boolean broke = false;
				for(int j=0;j<col.length;j++){
					
					for(int k=0;k<filters.length;k++){
						if(labels.get(j).equals(filters[k].column) && Pattern.matches(filters[k].regex, col[j])){
							data.remove(i);
							broke = true;
							break;
						}
					}
					
					if(broke){break;}
					
				}
			}
		}
		
		public void filterColumns(String[] validColumns){
			ArrayList<Integer> indicesToRemove = new ArrayList<Integer>();
			for(int i=(labels.size()-1);i>-1;i--){
				boolean found = false;
				for(int j=0;j<validColumns.length;j++){
					if(labels.get(i).equalsIgnoreCase(validColumns[j])){
						found = true;
						break;
					}
				}
				if(!found){
					labels.remove(i);
					indicesToRemove.add(i);
				}
			}
			
			for(int i=0;i<data.size();i++){
				ArrayList<String> cols = new ArrayList<String>();
				for(int j=0;j<data.get(i).length;j++){
					if(!indicesToRemove.contains(new Integer(j))){
						cols.add(data.get(i)[j]);
					}
				}
				data.set(i, cols.toArray(new String[cols.size()]));
			}
		}
		
		public void filterSapExport(SapExport sapExport, String columnName){
			ArrayList<String> numbers = new ArrayList<String>();
			
			for(int i=0;i<sapExport.data.size();i++){
				for(int j=0;j<sapExport.data.get(i).length;j++){
					if(sapExport.labels.get(j).equalsIgnoreCase(columnName)){
						numbers.add(sapExport.data.get(i)[j]);
					}
				}
			}
			
			for(int i=(data.size()-1); i>-1; i--){
				for(int j=0;j<data.get(i).length;j++){
					if(labels.get(j).equalsIgnoreCase("Material") && numbers.contains(data.get(i)[j])){
						data.remove(i);
						break;
					}
				}
			}
		}
		
		public Map<String, ArrayList<String[]>> seperateIntoSheetMap(Map<String, ArrayList<String[]>> sheetMap, String column){
			//iterate through lines
			for(int i=0;i<data.size();i++){
				//and cols
				if(data.get(i).length != labels.size()){
					//System.out.println(Arrays.toString(data.get(i))+ " doesnt have " + labels.size() + " cols");
					continue;
				}
				
				for(int j=0;j<data.get(i).length;j++){
					String label = labels.get(j).trim();
					//System.out.println(label);
					if(label.equalsIgnoreCase(column)){
						if(sheetMap.containsKey(data.get(i)[j])){
							sheetMap.get(data.get(i)[j]).add(data.get(i));
						}else{
							sheetMap.put(data.get(i)[j], new ArrayList<String[]>());
							sheetMap.get(data.get(i)[j]).add(data.get(i));
						}
					}
				}
			}
			
			return sheetMap;
		}
		
		public void removeDuplicates(){
			ArrayList<String> material = new ArrayList<String>();
			int materialIndex = -1, chargeIndex = -1;
			for(int j=0;j<labels.size();j++){
				if(labels.get(j).equalsIgnoreCase("Material")){
					materialIndex = j;
				}else if(labels.get(j).equalsIgnoreCase("Charge")){
					chargeIndex = j;
				}
			}
			
			for(int i=(data.size()-1);i>-1;i--){
				if(material.contains(data.get(i)[materialIndex])){
					//duplicate !
//					System.out.println("Duplicate " + data.get(i)[materialIndex] + ", Charge: '" + data.get(i)[chargeIndex] + "'");
					if(data.get(i)[chargeIndex].trim().length() == 0){
						//without charge => delete
//						System.out.println("removed");
						data.remove(i);
					}
				}else{
					material.add(data.get(i)[materialIndex]);
				}
			}
		}
		
		public void sort(){
			int nameIndex = -1;
			for(int j=0;j<labels.size();j++){
				if(labels.get(j).equalsIgnoreCase("Materialkurztext")){
					nameIndex = j;
				}
			}
			
			Collections.sort(data, compareRows(nameIndex));
		}
		
		private Comparator<String[]> compareRows(final int columnIndex){   
			return new Comparator<String[]>(){
				@Override
				public int compare(String[] r1, String[] r2){
					return r1[columnIndex].compareTo(r2[columnIndex]);
				}        
			};
		}  
	}
}
