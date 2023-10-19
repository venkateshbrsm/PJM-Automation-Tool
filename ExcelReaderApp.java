package ExcelReader;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.RowFilter;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.event.MouseInputAdapter;
import javax.swing.event.TableModelEvent;
import javax.swing.event.TableModelListener;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableRowSorter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReaderApp {
	private JFrame frame;
	private JPanel panel;
	private JButton openButton;
	private JTable table;
	private DefaultTableModel model;
	private JPanel footerPanel;
	private JButton saveToSharePointButton;
	private JButton saveButton;
	private JButton closeButton;
	private JTextField searchField;
	private JButton searchButton;
	private TableRowSorter<DefaultTableModel> sorter;
	private String filePath = null;
	private SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yy");
	JLabel pathToFile = null;
	MouseInputAdapter mouseAdapter = null;
	int lastClickedRow = -1;
	int lastClickedColumn = -1;

	public ExcelReaderApp() {
		frame = new JFrame("Release PJM - Automation Initiative Tools");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setSize(800, 600);
		panel = new JPanel();
		panel.setBackground(Color.LIGHT_GRAY);
		openButton = new JButton("Choose Jira Dump");
		pathToFile = new JLabel(filePath);
		model = new DefaultTableModel();
		table = new JTable();
		// Add a TableModelListener to the DefaultTableModel
		table.setRowHeight(table.getRowHeight() * 5); // Increase row height by 3 times
		// Create a MouseInputAdapter to handle mouse events at the cell level
		mouseAdapter = new MouseInputAdapter() {

			@Override
			public void mouseClicked(MouseEvent e) {
				// TODO Auto-generated method stub
				super.mouseClicked(e);
				lastClickedRow = table.rowAtPoint(e.getPoint());
				lastClickedColumn = table.columnAtPoint(e.getPoint());
				saveWithoutAlert(lastClickedRow, lastClickedColumn);
			}

		};

		// Add the MouseInputAdapter to the table
		table.addMouseListener(mouseAdapter);

		openButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				int returnValue = fileChooser.showOpenDialog(null);

				if (returnValue == JFileChooser.APPROVE_OPTION) {
					try {
						filePath = fileChooser.getSelectedFile().getAbsolutePath();
						model = readExcel(filePath);
						table.setModel(model);
						setTableColumnWrapping();
						addTableSorting();
						addSearchField();
						saveButton = new JButton("Save");
						closeButton = new JButton("Close");
						footerPanel.add(saveButton);
						saveToSharePointButton = new JButton("Save Excel to SharePoint");
						saveActionListener();
						footerPanel.add(saveToSharePointButton);
						footerPanel.add(closeButton);
						frame.setSize(1200, 800);
						setTableColumnProperties();
						sorter = new TableRowSorter<>(model);
						table.setRowSorter(sorter);
						closeButton.addActionListener(new ActionListener() {
							@Override
							public void actionPerformed(ActionEvent e) {
								frame.dispose();
							}
						});
					} catch (Exception ex) {
						ex.printStackTrace();
						JOptionPane.showMessageDialog(frame, "Error reading the Excel file: " + ex.getMessage());
					}
				}
			}
		});
		panel.add(openButton);
		frame.add(panel, BorderLayout.PAGE_START);
		frame.add(new JScrollPane(table), BorderLayout.CENTER);

		footerPanel = new JPanel();
		frame.add(footerPanel, BorderLayout.AFTER_LAST_LINE);
	}

	private void saveActionListener() {
		saveButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				save();
			}
		});
		saveButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				// Programmatically select a cell and trigger a click
				int row = 0; // Row index
				int column = 0; // Column index
				table.setRowSelectionInterval(row, row);
				table.setColumnSelectionInterval(column, column);
				table.changeSelection(row, column, false, false);
				table.requestFocus();
				table.editCellAt(row, column);
				table.transferFocus();

				// Perform some action as if the cell was clicked
				// For demonstration, we're just printing the selected cell value
				Object selectedValue = table.getValueAt(row, column);
				System.out.println("Clicked Cell Value: " + selectedValue);
			}
		});
	}

	private void save() {

		if (filePath != null && model.getColumnCount() > 0) {
			try {
				writeExcel();
				JOptionPane.showMessageDialog(frame, "Data saved successfully.");
			} catch (Exception ex) {
				ex.printStackTrace();
				JOptionPane.showMessageDialog(frame, "Error saving data to Excel: " + ex.getMessage());
			}
		} else {
			JOptionPane.showMessageDialog(frame, "No data to save or file not selected.");
		}
		// frame.dispose();
	}

	private void saveWithoutAlert(int row, int column) {

		if (filePath != null && model.getColumnCount() > 0) {
			try {
				writeExcel(row, column);
				// JOptionPane.showMessageDialog(frame, "Data saved successfully.");
			} catch (Exception ex) {
				ex.printStackTrace();
				JOptionPane.showMessageDialog(frame, "Error saving data to Excel: " + ex.getMessage());
			}
		} else {
			JOptionPane.showMessageDialog(frame, "No data to save or file not selected.");
		}

	}

	private DefaultTableModel readExcel(String filePath) {
		DefaultTableModel newModel = new DefaultTableModel();
		String age = "0";
		try (Workbook workbook = new XSSFWorkbook(filePath)) {
			Sheet sheet = workbook.getSheetAt(0);
			Row headerRow = sheet.getRow(0);

			for (Cell cell : headerRow) {
				newModel.addColumn(cell.getStringCellValue());
			}
			int colCount = newModel.getColumnCount();
			newModel.addColumn("Age");
			for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
				Row dataRow = sheet.getRow(rowIndex);
				Object[] rowData = new Object[newModel.getColumnCount() + 1];

				for (int columnIndex = 0; columnIndex < newModel.getColumnCount(); columnIndex++) {
					Cell cell = dataRow.getCell(columnIndex);
					try {
						age = String.valueOf(Integer.parseInt(getAge(getValue(cell, newModel.getColumnCount()))));
					} catch (Exception e) {
						// TODO: handle exception
					}
					if (cell != null)
						rowData[columnIndex] = getValue(cell, newModel.getColumnCount());
				}

				rowData[colCount] = age;
				newModel.addRow(rowData);

			}

		} catch (Exception e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(frame, "Error reading the Excel file: " + e.getMessage());
		}
		frame.setTitle("Release PJM - Automation Initiative Tools");
		frame.setTitle(frame.getTitle() + "       " + filePath);

		// saveToSharePointButton.addMouseListener(mouseAdapter);
		return newModel;
	}

	private String getAge(String value) {
		Date pastDate = null;
		Date currentDate = new Date();
		long differenceInMillis = 0;
		long differenceInDays = 0;
		try {
			if (dateFormat.parse(String.valueOf(value)) instanceof Date) {
				pastDate = dateFormat.parse(String.valueOf(value));
				differenceInMillis = currentDate.getTime() - pastDate.getTime();
				differenceInDays = differenceInMillis / (1000 * 60 * 60 * 24);
			}
		} catch (Exception e) {
			// TODO: handle exception
		}
		if (differenceInDays > 0)
			return String.valueOf(differenceInDays);
		else
			return value;
	}

	private String getValue(Cell cell, int columnIndex) {
		String data = null;
		if (cell.getCellType() == CellType.STRING)
			data = cell.getStringCellValue();
		else if (cell.getCellType() == CellType.NUMERIC && columnIndex != 4)
			data = String.valueOf(cell.getNumericCellValue());
		else {
			Instant instant = cell.getDateCellValue().toInstant();
			LocalDate localDate = instant.atZone(ZoneId.systemDefault()).toLocalDate();
			DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yy");
			data = localDate.format(formatter);
		}
		return data;
	}

	private void writeExcel(int row, int column) throws Exception {
		Object cellValue = table.getValueAt(row, column);
		model.setValueAt(cellValue, row, column);
		model.fireTableCellUpdated(row, column);
		createFreshSheet();
	}

	private void writeExcel() throws Exception {
		createFreshSheet();
	}

	private void createFreshSheet() throws Exception {
		try {
			FileInputStream inputStream = new FileInputStream(new File(filePath));
			Workbook workbook = WorkbookFactory.create(inputStream);
			Sheet sheet = workbook.getSheetAt(0);

			for (int rowIndex = 1; rowIndex <= model.getRowCount(); rowIndex++) {
				Row dataRow = sheet.createRow(rowIndex);
				for (int columnIndex = 0; columnIndex < model.getColumnCount(); columnIndex++) {
					Cell cell = dataRow.createCell(columnIndex);
					cell.setCellValue(table.getValueAt(rowIndex - 1, columnIndex).toString());
				}
			}

			FileOutputStream outputStream = new FileOutputStream(filePath);
			workbook.write(outputStream);
			outputStream.close();

		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}

	private void setTableColumnWrapping() {
		DefaultTableCellRenderer cellRenderer = new DefaultTableCellRenderer() {
			{
				setHorizontalAlignment(SwingConstants.CENTER); // Center text within cells
			}

			@Override
			public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected,
					boolean hasFocus, int row, int column) {
				JTextArea textArea = new JTextArea(value.toString());
				textArea.setLineWrap(true);
				textArea.setWrapStyleWord(true);
				textArea.setOpaque(true);
				textArea.setFont(getFont());
				return textArea;
			}
		};

		// Apply the cell renderer to all columns
		for (int i = 0; i < table.getColumnCount(); i++) {
			table.getColumnModel().getColumn(i).setCellRenderer(cellRenderer);
		}
	}

	private void addTableSorting() {
		table.setAutoCreateRowSorter(true);
	}

	private void addSearchField() {
		JPanel searchPanel = new JPanel();
		searchField = new JTextField();
		searchField.setColumns(20); // Increase the width by 20 times
		Insets margin = new Insets(10, 50, 10, 20);
		searchField.setMargin(margin);
		searchButton = new JButton("Search");
		searchButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				String text = searchField.getText();
				if (text.trim().length() == 0) {
					sorter.setRowFilter(null);
				} else {
					sorter.setRowFilter(RowFilter.regexFilter("(?i)" + text));
				}
			}
		});

		searchPanel.add(searchField);
		searchPanel.add(searchButton);
		panel.add(searchPanel);
	}

	private void setTableColumnProperties() {
		DefaultTableCellRenderer cellRenderer = new DefaultTableCellRenderer() {
			{
				setHorizontalAlignment(SwingConstants.CENTER); // Center text within cells
			}

			@Override
			public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected,
					boolean hasFocus, int row, int column) {
				return super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
			}
		};

		// Apply the cell renderer to all columns
		for (int i = 0; i < table.getColumnCount(); i++) {
			table.getColumnModel().getColumn(i).setCellRenderer(cellRenderer);
		}
	}

	public static Object getCellValueByColumnName(JTable table, String columnName, int rowIndex) {
		int columnCount = table.getColumnCount();
		for (int i = 0; i < columnCount; i++) {
			if (table.getColumnName(i).equals(columnName)) {
				return table.getValueAt(rowIndex, i);
			}
		}
		return null; // Column with the given name not found
	}

	public void display() {
		frame.setVisible(true);
	}

	public static void main(String[] args) {
		SwingUtilities.invokeLater(() -> {
			ExcelReaderApp app = new ExcelReaderApp();
			app.display();

		});
	}
}
