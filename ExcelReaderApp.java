package ExcelReader;

import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.RowFilter;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
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
	private JTextField searchField;
	private JButton searchButton;
	private TableRowSorter<DefaultTableModel> sorter;
	private String filePath = null;

	public ExcelReaderApp() {
		frame = new JFrame("Release PJM - Automation Initiative Tools");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setSize(800, 600);

		panel = new JPanel();
		openButton = new JButton("Open Excel File");
		table = new JTable();
		model = new DefaultTableModel();
		table.setRowHeight(table.getRowHeight() * 3); // Increase row height by 3 times
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
						saveToSharePointButton = new JButton("Save Excel to SharePoint");
						saveButton = new JButton("Save");
						saveActionListener();
						footerPanel.add(saveToSharePointButton);
						footerPanel.add(saveButton);
						frame.setSize(1200, 800);
						setTableColumnProperties();
						sorter = new TableRowSorter<>(model);
						table.setRowSorter(sorter);

					} catch (Exception ex) {
						ex.printStackTrace();
						JOptionPane.showMessageDialog(frame, "Error reading the Excel file: " + ex.getMessage());
					}
				}
			}
		});

		panel.add(openButton);
		frame.add(panel, BorderLayout.NORTH);
		frame.add(new JScrollPane(table), BorderLayout.CENTER);
		footerPanel = new JPanel();
		frame.add(footerPanel, BorderLayout.AFTER_LAST_LINE);
	}

	private void saveActionListener() {
	    saveButton.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
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
	        }
	    });
	}

	private DefaultTableModel readExcel(String filePath) {
		DefaultTableModel newModel = new DefaultTableModel();

		try (Workbook workbook = new XSSFWorkbook(filePath)) {
			Sheet sheet = workbook.getSheetAt(0);
			Row headerRow = sheet.getRow(0);

			for (Cell cell : headerRow) {
				newModel.addColumn(cell.getStringCellValue());
			}

			for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
				Row dataRow = sheet.getRow(rowIndex);
				Object[] rowData = new Object[newModel.getColumnCount()];

				for (int columnIndex = 0; columnIndex < newModel.getColumnCount(); columnIndex++) {
					Cell cell = dataRow.getCell(columnIndex);
					rowData[columnIndex] = getValue(cell, newModel.getColumnCount());
				}

				newModel.addRow(rowData);

			}
		} catch (Exception e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(frame, "Error reading the Excel file: " + e.getMessage());
		}

		return newModel;
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

	private void writeExcel() throws Exception {
		// removeSheet();
		createFreshSheet();
	}

	private void createFreshSheet() throws Exception {
		try {
			FileInputStream inputStream = new FileInputStream(new File(filePath));
			Workbook workbook = WorkbookFactory.create(inputStream);
			Sheet sheet = workbook.getSheetAt(0);
			/*
			 * for (int rowIndex = 0; rowIndex < model.getRowCount(); rowIndex++) {
			 * sheet.removeRow(sheet.getRow(rowIndex)); }
			 */
			for (int rowIndex = 1; rowIndex <= model.getRowCount(); rowIndex++) {
				Row dataRow = sheet.createRow(rowIndex);
				for (int columnIndex = 0; columnIndex < model.getColumnCount(); columnIndex++) {
					Cell cell = dataRow.createCell(columnIndex);
					cell.setCellValue(model.getValueAt(rowIndex - 1, columnIndex).toString());
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
