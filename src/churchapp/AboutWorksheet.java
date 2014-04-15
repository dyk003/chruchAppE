/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package churchapp;

/**
 *
 * @author David
 */
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.print.PrinterException;
import java.io.File;
import java.io.IOException;
import java.text.MessageFormat;
import java.util.Locale;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.MediaName;
import javax.print.attribute.standard.OrientationRequested;
import javax.print.attribute.standard.RequestingUserName;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenuItem;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumnModel;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
 
public class AboutWorksheet {
    static JTable listTable;
    static int trackNumber;
    Workbook workbook;  
    private JFrame frame; 
    Sheet sheet;
    Cell[] district, group, koreanName, firstName, lastName, title, spouse;
    Cell[] child, hphone, cphone, street, city, state, zip, status; 
    JTextField tFindName, tKoreanName, tFirstName, tLastName, tTitle;  
    JTextField tSpouse, tChild, tDistrict, tGroup, tPhoneH, tPhoneC; 
    JTextField tStreet, tCity, tState, tZip;
    JCheckBox cStatus;
    JLabel notFound;
    JButton bPrevious, bNext;
    JMenuItem mEdit;
    String findKoreanName, findFirstName, findLastName;

    public AboutWorksheet(Workbook workbook, JTextField tFindName, 
        JTextField tKoreanName, JTextField tFirstName, JTextField tLastName, 
        JTextField tTitle, JTextField tSpouse, JTextField tChild, 
        JTextField tDistrict, JTextField tGroup, JTextField tPhoneH, 
        JTextField tPhoneC, JTextField tStreet, JTextField tCity, 
        JTextField tState, JTextField tZip, JCheckBox cStatus,
        String findKoreanName, String findFirstName, String findLastName,
        JLabel notFound, JButton bPrevious, JButton bNext, JMenuItem mEdit) { 
            this.workbook = workbook;   
            this.findKoreanName = findKoreanName;
            this.findFirstName = findFirstName;
            this.findLastName = findLastName;
            this.tKoreanName = tKoreanName;
            this.tFirstName = tFirstName;
            this.tLastName = tLastName;
            this.tTitle = tTitle;
            this.tSpouse = tSpouse;
            this.tChild = tChild;
            this.tDistrict = tDistrict;
            this.tGroup = tGroup;
            this.tPhoneH = tPhoneH;
            this.tPhoneC = tPhoneC;
            this.tStreet = tStreet;
            this.tCity = tCity;
            this.tState = tState;
            this.tZip = tZip;
            this.cStatus = cStatus;
            this.notFound = notFound;
            this.bPrevious = bPrevious;
            this.bNext = bNext;
            this.mEdit = mEdit;
            sheet = workbook.getSheet(0); 
            district = sheet.getColumn(0);
            group = sheet.getColumn(1);
            koreanName = sheet.getColumn(2);
            firstName = sheet.getColumn(3);
            lastName = sheet.getColumn(4);
            title = sheet.getColumn(5);
            spouse = sheet.getColumn(6);  
            child = sheet.getColumn(7);  
            hphone = sheet.getColumn(8);
            cphone = sheet.getColumn(9);  
            street = sheet.getColumn(10);
            city = sheet.getColumn(11);
            state = sheet.getColumn(12);
            zip = sheet.getColumn(13);   
            status = sheet.getColumn(14); 
    }
	
    public void getInfoByKoreanName() {
        for (int i = 0; i < state.length; i ++) { 
            if (koreanName[i].getContents().trim().toLowerCase().contains(findKoreanName.trim().toLowerCase())
                || spouse[i].getContents().trim().toLowerCase().contains(findKoreanName.trim().toLowerCase())
                || child[i].getContents().trim().toLowerCase().contains(findKoreanName.trim().toLowerCase())) { 
                trackNumber = i;
                displayInfo(i);
                bPrevious.setEnabled(true);
                bNext.setEnabled(true);
                mEdit.setEnabled(true);
                break; 
            } else {
                notFound.setForeground(Color.RED);
                notFound.setFont(notFound.getFont().deriveFont (12.0f)); 
                notFound.setText("Not Found / 입력한 항목이 없습니다.");
            }
        }
    }
    
    public void getInfoByName() { 
        for (int i = 0; i < state.length; i ++) { 
            if (firstName[i].getContents().trim().toLowerCase().contains(findFirstName.trim().toLowerCase()) 
                && lastName[i].getContents().trim().toLowerCase().contains(findLastName.trim().toLowerCase())
                || spouse[i].getContents().trim().toLowerCase().contains(findFirstName.trim().toLowerCase()) 
                || child[i].getContents().trim().toLowerCase().contains(findFirstName.trim().toLowerCase())) {
                trackNumber = i;
                displayInfo(i);
                bPrevious.setEnabled(true);
                bNext.setEnabled(true);
                mEdit.setEnabled(true);
                break; 
            } else {
                notFound.setForeground(Color.RED);
                notFound.setFont(notFound.getFont().deriveFont(12.0f)); 
                notFound.setText("Not Found / 입력한 항목이 없습니다.");
            } 
        }
    }
    
    public void update() throws IOException, WriteException { 
        WritableWorkbook wworkbook = Workbook.createWorkbook(new File("roster.xls"), workbook);
        WritableSheet wsheet = wworkbook.getSheet(0);   
        
        WritableCell cellDistrict = new Label(0, trackNumber, tDistrict.getText()); 
        WritableCell cellGroup = new Label(1, trackNumber, tGroup.getText());
        WritableCell cellKoreanName = new Label(2, trackNumber, tKoreanName.getText());
        WritableCell cellFirstName = new Label(3, trackNumber, tFirstName.getText());
        WritableCell cellLastName = new Label(4, trackNumber, tLastName.getText());
        WritableCell cellTitle = new Label(5, trackNumber, tTitle.getText());
        WritableCell cellSpouse = new Label(6, trackNumber, tSpouse.getText());
        WritableCell cellChild = new Label(7, trackNumber, tChild.getText());
        WritableCell cellPhoneH = new Label(8, trackNumber, tPhoneH.getText());
        WritableCell cellPhoneC = new Label(9, trackNumber, tPhoneC.getText());
        WritableCell cellStreet = new Label(10, trackNumber, tStreet.getText());
        WritableCell cellCity = new Label(11, trackNumber, tCity.getText());
        WritableCell cellState = new Label(12, trackNumber, tState.getText());
        WritableCell cellZip = new Label(13, trackNumber, tZip.getText()); 
        WritableCell cellStatus;
        if (cStatus.isSelected()) {
            cellStatus = new Label(14, trackNumber, "X");    
        } else {
            cellStatus = new Label(14, trackNumber, "");   
        }
        
        wsheet.addCell(cellDistrict);
        wsheet.addCell(cellGroup);
        wsheet.addCell(cellKoreanName);
        wsheet.addCell(cellFirstName);
        wsheet.addCell(cellLastName);
        wsheet.addCell(cellTitle);
        wsheet.addCell(cellSpouse);
        wsheet.addCell(cellChild);
        wsheet.addCell(cellPhoneH);
        wsheet.addCell(cellPhoneC);
        wsheet.addCell(cellStreet);
        wsheet.addCell(cellCity);
        wsheet.addCell(cellState);
        wsheet.addCell(cellZip);
        wsheet.addCell(cellStatus); 
        
        wworkbook.write();
        wworkbook.close(); 
    }
    
    public void getNext() {  
        if (trackNumber != state.length-1) {
            trackNumber++;
            displayInfo(trackNumber);  
        }    
    }
    
    public void getPrevious() {  
        if (trackNumber != 0) {
            trackNumber--;
            displayInfo(trackNumber);   
        }
    }
     
    public void getDistrict(int dst, int grp) {
        frame = new SizedFrame();
        if (dst == 50) {
            frame.setTitle("Tacoma New Life Church / 타코마새생명장로교회 EM 교구");
        } else {
            frame.setTitle("Tacoma New Life Church / 타코마새생명장로교회 제" + dst + "교구");
        }
        
        frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        frame.setVisible(true); 
 
        DefaultTableModel table = setupMenu();

        for (int j = 1; j <= grp; j++) {
            for (int i = 1; i < state.length; i++) {  
                if (dst == 50) {
                    if (district[i].getContents().trim().equals("EM")) {
                        if (group[i].getContents().trim().equals(String.valueOf(j))) {
                            displayItems(i, table);
                        }
                    }
                } else {
                    if (district[i].getContents().trim().equals(String.valueOf(dst))) {
                        if (group[i].getContents().trim().equals(String.valueOf(j))) {
                            displayItems(i, table);
                        }
                    }
                }
            }
        } 
         
        setupFrame(table, 1400, 730, "", dst, 0);
    }
    
    public void getGroup(int dst, int grp) {
        frame = new JFrame();
        if (dst == 50) {
            frame.setTitle("Tacoma New Life Church / 타코마새생명장로교회 EM 교구 " + grp + "구역");
        } else {
            frame.setTitle("Tacoma New Life Church / 타코마새생명장로교회 제" + dst
                    + "교구 " + grp + "구역");
        }
        frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);   
        frame.setVisible(true);  
 
        DefaultTableModel table = setupMenu();

        for (int i = 1; i < state.length; i++) {  
            if (dst == 50) {
                if (district[i].getContents().trim().equals("EM")) {
                    if (group[i].getContents().trim().equals(String.valueOf(grp))) {
                        displayItems(i, table);
                    }
                } 
            } else {
                if (district[i].getContents().trim().equals(String.valueOf(dst))) {
                    if (group[i].getContents().trim().equals(String.valueOf(grp))) {
                        displayItems(i, table); 
                    }
                }
            }
        }

        setupFrame(table, 1200, 400, "", dst, grp); 
    } 
    
    public void getList() {
        frame = new SizedFrame();
        frame.setTitle("Tacoma New Life Church / 타코마새생명장로교회 전체 명단");
        frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        frame.setVisible(true);
//		printSetup printJob = new printSetup();
//		frame.setContentPane(panel);  
//		panel.setBackground(Color.WHITE); 


        DefaultTableModel table = setupMenu();

        for (int i = 1; i < state.length; i++) {  
            displayItems(i, table);
        }

        setupFrame(table, 1400, 730, "전체 교인", 0, 0);
    } 
    
    public void getTitleGroup(String ttitle) {
        frame = new JFrame();
        frame.setTitle("Tacoma New Life Church / 타코마새생명장로교회 (" + title + ")");
        
        frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);   
        frame.setVisible(true);  
 
        DefaultTableModel table = setupMenu();

        for (int i = 1; i < state.length; i++) {   
            if (title[i].getContents().trim().equals(ttitle)) { 
                displayItems(i, table);  
            } 
        }
        
        if (ttitle.equals("목사") || ttitle.equals("전도사")) {
            setupFrame(table, 1300, 400, ttitle, 0, 0);
        } else if (ttitle.equals("장로")) {
            setupFrame(table, 1400, 600, ttitle, 0, 0); 
        } else {
            setupFrame(table, 1400, 730, ttitle, 0, 0);
        }
    }
    
    private void setupFrame(DefaultTableModel table, int x, int y, String ttitle, int dst, int grp) {
        setupTable(table); 
        JPanel buttonPanel = new JPanel();
        JScrollPane pane = new JScrollPane(listTable); 
        JButton printButton = setupPrint(ttitle, dst, grp);
        buttonPanel.add(printButton); 
        frame.setSize(x, y); 
        frame.add(pane, BorderLayout.CENTER); 
        frame.add(buttonPanel, BorderLayout.SOUTH); 
    }
    

    private JButton setupPrint(String ttitle, int dst, int grp) {
        final MessageFormat footerFormat;
        if (dst != 0 && grp != 0 && dst != 50) {
            //final MessageFormat headerFormat= new MessageFormat("타코마새생명장로교회");
            footerFormat = new  MessageFormat(dst + "교구" 
                    + grp + "구역 목록  타코마새생명장로교회  페이지{0}"); 
        } else if (dst != 0 && grp == 0 && dst != 50) { 
            //final MessageFormat headerFormat= new MessageFormat("타코마새생명장로교회");
            footerFormat = new  MessageFormat(dst + "구역 전체 목록  타코마새생명장로교회  페이지{0}"); 
        } else if (dst == 0 && grp == 0 && ttitle.equals("")) { 
            //final MessageFormat headerFormat= new MessageFormat("타코마새생명장로교회");
            footerFormat = new  MessageFormat("전 교인 목록  타코마새생명장로교회  페이지{0}"); 
        } else if (dst == 50 && grp == 0) {
            //final MessageFormat headerFormat= new MessageFormat("타코마새생명장로교회");
            footerFormat = new  MessageFormat("EM교구 전체 목록  타코마새생명장로교회  페이지{0}"); 
        }else if (dst ==50 && grp != 0) {
            //final MessageFormat headerFormat= new MessageFormat("타코마새생명장로교회");
            footerFormat = new  MessageFormat("EM교구" + grp + "구역 목록  타코마새생명장로교회  페이지{0}");
        } else {
            //final MessageFormat headerFormat= new MessageFormat("타코마새생명장로교회");
            footerFormat = new  MessageFormat(ttitle + " 목록  타코마새생명장로교회  페이지{0}");
        }
        JButton printButton = new JButton("Print / 인쇄하기"); 

        printButton.addActionListener((new ActionListener() {
            public void actionPerformed(final ActionEvent the_event) {  
                try {
                    PrintRequestAttributeSet aset = new HashPrintRequestAttributeSet();
                    aset.add(OrientationRequested.LANDSCAPE);
                    aset.add(MediaName.ISO_A4_WHITE);
                    aset.add(new RequestingUserName("ichebyki", Locale.US));
                    listTable.print(JTable.PrintMode.FIT_WIDTH, null, footerFormat, true, aset, true );
                } catch (PrinterException e) { 
                    e.printStackTrace();
                }
            }
        }));
        return printButton;
    }
    
    private void setupTable(DefaultTableModel table) {
        listTable = new JTable(table);
        listTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF); 
        listTable.setAutoCreateRowSorter(true);   
        resizeColumnWidth(listTable);
        listTable.setIntercellSpacing(new Dimension(10,0));
//	    listTable.setRowHeight(19);
        listTable.getTableHeader().setFont(new Font("SansSerif", Font.ITALIC, 12)); 
    }
    private void displayInfo(int i) {
        notFound.setText("");
        tDistrict.setText(district[i].getContents().trim());
        tGroup.setText(group[i].getContents().trim()); 
        tKoreanName.setText(koreanName[i].getContents().trim());  
        tFirstName.setText(firstName[i].getContents().trim());
        tLastName.setText(lastName[i].getContents().trim());
        tTitle.setText(title[i].getContents().trim());
        tSpouse.setText(spouse[i].getContents().trim());
        tChild.setText(child[i].getContents().trim());
        tPhoneH.setText(hphone[i].getContents().trim());
        tPhoneC.setText(cphone[i].getContents().trim());
        tStreet.setText(street[i].getContents().trim());
        tCity.setText(city[i].getContents().trim());
        tState.setText(state[i].getContents().trim());
        tZip.setText(zip[i].getContents().trim()); 
        if (status[i].getContents().trim().equals("X")) {
            cStatus.setSelected(true);
        } else {
            cStatus.setSelected(false);
        }
    }
    
    private void resizeColumnWidth(JTable listTable) {
        final TableColumnModel columnModel = listTable.getColumnModel();
        for (int column = 0; column < listTable.getColumnCount(); column++) {
            int width = 40; // Min width
            for (int row = 0; row < listTable.getRowCount(); row++) {
                TableCellRenderer renderer = listTable.getCellRenderer(row, column);
                Component comp = listTable.prepareRenderer(renderer, row, column);
                width = Math.max(comp.getPreferredSize().width, width);
            }
            columnModel.getColumn(column).setPreferredWidth(width+10);
        }
    }
    
    private DefaultTableModel setupMenu() {
        Object[] columnNames = {"교구", "구역", "이름", "First Name", 
            "Last Name", "직분", "배우자", "자녀", "집전화", "핸드폰전화", "Street", 
            "City", "Zip", "비활동"};
        Object[][] rowData = {};  
        DefaultTableModel table = new DefaultTableModel(rowData, columnNames);
        return table;
    }
    
    private void displayItems(int i, DefaultTableModel table) { 
        table.addRow(new Object[] {district[i].getContents(), 
        group[i].getContents(),
        koreanName[i].getContents(), firstName[i].getContents(),
        lastName[i].getContents(), title[i].getContents(), 
        spouse[i].getContents(), child[i].getContents(), 
        hphone[i].getContents(),cphone[i].getContents(), 
        street[i].getContents(), city[i].getContents(), 
        zip[i].getContents(),status[i].getContents()}); 
    }
     
    public void saveList() throws RowsExceededException, WriteException, IOException {
        WritableWorkbook wworkbook = Workbook.createWorkbook(new File("roster.xls"));
        WritableSheet wsheet = wworkbook.createSheet("First Sheet", 0);  
        for (int i = 0; i < state.length; i++) {
            wsheet.addCell(new Label(0, i, district[i].getContents()));
            wsheet.addCell(new Label(1, i, group[i].getContents()));
            wsheet.addCell(new Label(2, i, koreanName[i].getContents())); 
            wsheet.addCell(new Label(3, i, firstName[i].getContents()));
            wsheet.addCell(new Label(4, i, lastName[i].getContents()));
            wsheet.addCell(new Label(5, i, title[i].getContents()));
            wsheet.addCell(new Label(6, i, spouse[i].getContents()));
            wsheet.addCell(new Label(7, i, child[i].getContents()));
            wsheet.addCell(new Label(8, i, hphone[i].getContents()));
            wsheet.addCell(new Label(9, i, cphone[i].getContents()));
            wsheet.addCell(new Label(10, i, street[i].getContents()));
            wsheet.addCell(new Label(11, i, city[i].getContents()));
            wsheet.addCell(new Label(12, i, state[i].getContents()));
            wsheet.addCell(new Label(13, i, zip[i].getContents()));
            wsheet.addCell(new Label(14, i, status[i].getContents()));
        } 
        wworkbook.write();
        wworkbook.close(); 
    } 
}

class SizedFrame extends JFrame { 
	
	private static final long serialVersionUID = 1L;

	public SizedFrame() {
		Toolkit kit = Toolkit.getDefaultToolkit();
		Dimension screenSize = kit.getScreenSize();  
		setSize(screenSize.width, screenSize.height-40);
		setLocationByPlatform(false); 
	} 
} 
