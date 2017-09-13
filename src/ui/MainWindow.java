package ui;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.GridLayout;
import javax.swing.JTextField;

import core.FileProcessor;
import core.EntryPoint;

import java.awt.GridBagLayout;
import java.awt.GridBagConstraints;
import java.awt.Insets;
import javax.swing.JLabel;
import javax.swing.DefaultComboBoxModel;
import javax.swing.DefaultListModel;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JList;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.awt.event.ActionEvent;
import javax.swing.SwingConstants;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.JComboBox;

public class MainWindow {

	public JFrame frmExcelDocumentSorter;
	private JTextField textFieldFullPathToOutputFile;
	protected static FileProcessor fileProcessor;
	final JFileChooser openFileChooser;
    final JFileChooser saveFileChooser;
	
	/**
	 * Create the application.
	 */
	public MainWindow() {
		openFileChooser = new JFileChooser();
		saveFileChooser = new JFileChooser();
		openFileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files","xls", "xlsx"));
		fileProcessor = new FileProcessor();
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmExcelDocumentSorter = new JFrame();
		frmExcelDocumentSorter.setTitle("Excel Document Sorter");
		frmExcelDocumentSorter.setBounds(100, 100, 496, 280);
		frmExcelDocumentSorter.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		JPanel panel = new JPanel();
		frmExcelDocumentSorter.getContentPane().add(panel, BorderLayout.CENTER);
		panel.setMinimumSize(new Dimension(0,0));
		GridBagLayout gbl_panel = new GridBagLayout();
		gbl_panel.columnWidths = new int[] {51, 65, 60, 30, 31, 200, 51};
		gbl_panel.rowHeights = new int[] {30, 20, 20, 20, 35, 23, 30, 0};
		gbl_panel.columnWeights = new double[]{1.0, 1.0, 0.0, 0.0, 0.0, 0.0, 1.0};
		gbl_panel.rowWeights = new double[]{1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0};
		panel.setLayout(gbl_panel);
		
		JPanel leftFillerPanel = new JPanel();
		GridBagConstraints gbc_leftFillerPanel = new GridBagConstraints();
		gbc_leftFillerPanel.gridheight = 7;
		gbc_leftFillerPanel.insets = new Insets(0, 0, 5, 5);
		gbc_leftFillerPanel.fill = GridBagConstraints.BOTH;
		gbc_leftFillerPanel.gridx = 0;
		gbc_leftFillerPanel.gridy = 0;
		panel.add(leftFillerPanel, gbc_leftFillerPanel);
		leftFillerPanel.setMinimumSize(new Dimension(25, 0));
		
		JPanel topFillerPanel = new JPanel();
		GridBagConstraints gbc_topFillerPanel = new GridBagConstraints();
		gbc_topFillerPanel.gridwidth = 5;
		gbc_topFillerPanel.insets = new Insets(0, 0, 5, 5);
		gbc_topFillerPanel.fill = GridBagConstraints.BOTH;
		gbc_topFillerPanel.gridx = 1;
		gbc_topFillerPanel.gridy = 0;
		panel.add(topFillerPanel, gbc_topFillerPanel);
		GridBagLayout gbl_topFillerPanel = new GridBagLayout();
		gbl_topFillerPanel.columnWidths = new int[] {212};
		gbl_topFillerPanel.rowHeights = new int[] {14};
		gbl_topFillerPanel.columnWeights = new double[]{0.0};
		gbl_topFillerPanel.rowWeights = new double[]{0.0};
		topFillerPanel.setLayout(gbl_topFillerPanel);
		
		JLabel lblErrorMessage = new JLabel("Error retrieving that file.");
		lblErrorMessage.setHorizontalAlignment(SwingConstants.CENTER);
		GridBagConstraints gbc_lblErrorMessage = new GridBagConstraints();
		gbc_lblErrorMessage.anchor = GridBagConstraints.NORTH;
		gbc_lblErrorMessage.gridx = 0;
		gbc_lblErrorMessage.gridy = 0;
		topFillerPanel.add(lblErrorMessage, gbc_lblErrorMessage);
		lblErrorMessage.setForeground(Color.red); 
		lblErrorMessage.setVisible(false);
		
		JPanel rightFillerPanel = new JPanel();
		GridBagConstraints gbc_rightFillerPanel = new GridBagConstraints();
		gbc_rightFillerPanel.insets = new Insets(0, 0, 5, 0);
		gbc_rightFillerPanel.gridheight = 7;
		gbc_rightFillerPanel.fill = GridBagConstraints.BOTH;
		gbc_rightFillerPanel.gridx = 6;
		gbc_rightFillerPanel.gridy = 0;
		panel.add(rightFillerPanel, gbc_rightFillerPanel);
		
		JTextField textFieldFullPathToInputFile = new JTextField("");
		GridBagConstraints gbc_textFieldFullPathToInputFile = new GridBagConstraints();
		gbc_textFieldFullPathToInputFile.fill = GridBagConstraints.BOTH;
		gbc_textFieldFullPathToInputFile.gridwidth = 2;
		gbc_textFieldFullPathToInputFile.anchor = GridBagConstraints.WEST;
		gbc_textFieldFullPathToInputFile.insets = new Insets(0, 0, 5, 5);
		gbc_textFieldFullPathToInputFile.gridx = 4;
		gbc_textFieldFullPathToInputFile.gridy = 1;
		panel.add(textFieldFullPathToInputFile, gbc_textFieldFullPathToInputFile);
		textFieldFullPathToInputFile.setVisible(false);
		textFieldFullPathToInputFile.setEditable(false);
		
		JLabel lblNewLabel = new JLabel("Sort By");
		GridBagConstraints gbc_lblNewLabel = new GridBagConstraints();
		gbc_lblNewLabel.gridwidth = 3;
		gbc_lblNewLabel.anchor = GridBagConstraints.SOUTHWEST;
		gbc_lblNewLabel.insets = new Insets(0, 0, 5, 5);
		gbc_lblNewLabel.gridx = 1;
		gbc_lblNewLabel.gridy = 2;
		panel.add(lblNewLabel, gbc_lblNewLabel);
		
		DefaultComboBoxModel<String> sortByListModel = new DefaultComboBoxModel<String>();
		JComboBox<String> sortByList = new JComboBox<String>(sortByListModel);
		GridBagConstraints gbc_sortByList = new GridBagConstraints();
		gbc_sortByList.gridwidth = 2;
		gbc_sortByList.fill = GridBagConstraints.BOTH;
		gbc_sortByList.insets = new Insets(0, 0, 5, 5);
		gbc_sortByList.gridx = 4;
		gbc_sortByList.gridy = 2;
		panel.add(sortByList, gbc_sortByList);
		
		JLabel lblOutputLocationoptional = new JLabel("Output Location (Optional)");
		GridBagConstraints gbc_lblOutputLocationoptional = new GridBagConstraints();
		gbc_lblOutputLocationoptional.fill = GridBagConstraints.HORIZONTAL;
		gbc_lblOutputLocationoptional.insets = new Insets(0, 0, 5, 5);
		gbc_lblOutputLocationoptional.gridwidth = 3;
		gbc_lblOutputLocationoptional.gridx = 1;
		gbc_lblOutputLocationoptional.gridy = 3;
		panel.add(lblOutputLocationoptional, gbc_lblOutputLocationoptional);
		
		textFieldFullPathToOutputFile = new JTextField();
		textFieldFullPathToOutputFile.setText(" ");
		textFieldFullPathToOutputFile.setColumns(10);
		GridBagConstraints gbc_textFieldFullPathToOutputFile = new GridBagConstraints();
		gbc_textFieldFullPathToOutputFile.gridwidth = 2;
		gbc_textFieldFullPathToOutputFile.anchor = GridBagConstraints.NORTH;
		gbc_textFieldFullPathToOutputFile.fill = GridBagConstraints.HORIZONTAL;
		gbc_textFieldFullPathToOutputFile.insets = new Insets(0, 0, 5, 5);
		gbc_textFieldFullPathToOutputFile.gridx = 4;
		gbc_textFieldFullPathToOutputFile.gridy = 3;
		panel.add(textFieldFullPathToOutputFile, gbc_textFieldFullPathToOutputFile);
		textFieldFullPathToOutputFile.setEditable(false);
		textFieldFullPathToOutputFile.setVisible(false);
		
		
		JButton btnFetchButton = new JButton("Get File");
		GridBagConstraints gbc_btnFetchButton = new GridBagConstraints();
		gbc_btnFetchButton.insets = new Insets(0, 0, 5, 5);
		gbc_btnFetchButton.gridwidth = 5;
		gbc_btnFetchButton.gridx = 1;
		gbc_btnFetchButton.gridy = 5;
		panel.add(btnFetchButton, gbc_btnFetchButton);
		btnFetchButton.setVisible(true);		
		
		JButton btnResetButton = new JButton("Reset");
		GridBagConstraints gbc_btnResetButton = new GridBagConstraints();
		gbc_btnResetButton.insets = new Insets(0, 0, 5, 5);
		gbc_btnResetButton.gridwidth = 5;
		gbc_btnResetButton.gridx = 1;
		gbc_btnResetButton.gridy = 5;
		panel.add(btnResetButton, gbc_btnResetButton);
		btnResetButton.setVisible(false);		
		
		JButton btnSortButton = new JButton("Sort");
		GridBagConstraints gbc_btnSortButton = new GridBagConstraints();
		gbc_btnSortButton.insets = new Insets(0, 0, 5, 5);
		gbc_btnSortButton.gridwidth = 5;
		gbc_btnSortButton.gridx = 1;
		gbc_btnSortButton.gridy = 5;
		panel.add(btnSortButton, gbc_btnSortButton);
		btnSortButton.setVisible(false);
		
		
		JPanel bottomFillerPanel = new JPanel();
		GridBagConstraints gbc_bottomFillerPanel = new GridBagConstraints();
		gbc_bottomFillerPanel.gridwidth = 5;
		gbc_bottomFillerPanel.insets = new Insets(0, 0, 5, 5);
		gbc_bottomFillerPanel.fill = GridBagConstraints.BOTH;
		gbc_bottomFillerPanel.gridx = 1;
		gbc_bottomFillerPanel.gridy = 6;
		panel.add(bottomFillerPanel, gbc_bottomFillerPanel);
		
		JButton btnPickFile = new JButton("Select Excel File");
		GridBagConstraints gbc_btnPickFile = new GridBagConstraints();
		gbc_btnPickFile.anchor = GridBagConstraints.WEST;
		gbc_btnPickFile.insets = new Insets(0, 0, 5, 5);
		gbc_btnPickFile.gridx = 4;
		gbc_btnPickFile.gridy = 1;
		panel.add(btnPickFile, gbc_btnPickFile);
		
		JButton btnOutputButton = new JButton("Select Output Location");
		GridBagConstraints gbc_btnOutputButton = new GridBagConstraints();
		gbc_btnOutputButton.anchor = GridBagConstraints.WEST;
		gbc_btnOutputButton.insets = new Insets(0, 0, 5, 5);
		gbc_btnOutputButton.gridx = 4;
		gbc_btnOutputButton.gridy = 3;
		panel.add(btnOutputButton, gbc_btnOutputButton);
		
		
		btnFetchButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent event) {
				try {
					//if the file fetch is successful set forward
					if(fileProcessor.fetchFile(textFieldFullPathToInputFile.getText())==true){
						btnFetchButton.setVisible(false);
						btnSortButton.setVisible(true);
						lblErrorMessage.setVisible(false);
						
						ArrayList<String> headersList = fileProcessor.getSheetColumnHeaders();
						for(int i=0;i<headersList.size();i++){
							sortByListModel.addElement(headersList.get(i));
						}
					}
				  //if it isn't successful show an error
				} catch (IOException ioException) {
					lblErrorMessage.setForeground(Color.red); 
					lblErrorMessage.setText("Error retrieving that file.");
					lblErrorMessage.setVisible(true);
				}
			}
		});
		
		btnSortButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent event) {
				if (fileProcessor.processFile(sortByList.getSelectedIndex())){
					
					//save file to disk
					//fileProcessor.saveToFile()
					
					//Reset
					btnSortButton.setVisible(false);
					btnFetchButton.setVisible(true);
					btnPickFile.setVisible(true);
					textFieldFullPathToInputFile.setText(" ");
					textFieldFullPathToInputFile.setVisible(false);
					sortByListModel.removeAllElements();
					
					//Notify user you are done
					lblErrorMessage.setForeground(Color.black);
					lblErrorMessage.setText("Done");
					lblErrorMessage.setVisible(true);
				} else {
					lblErrorMessage.setText("There was an error processing that file. What was successful has been placed in the folde with the original.");
				}
				
			}
		});
		
		btnPickFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent event) {
				//Hide any message already seen
				lblErrorMessage.setVisible(false);
				
				//open the file selection dialogue
				int foundFile = openFileChooser.showOpenDialog(btnPickFile);
				//if one is selected, show it in the text box
				if(foundFile == JFileChooser.APPROVE_OPTION ) {
					File selectedFile = openFileChooser.getSelectedFile();
					fileProcessor.setFile(selectedFile);
					
					btnPickFile.setVisible(false);
					textFieldFullPathToInputFile.setText(selectedFile.getPath());
					textFieldFullPathToInputFile.setVisible(true);
				}

			}
		});
		
		btnOutputButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent event) {
				//Hide any message already seen
				lblErrorMessage.setVisible(false);
				
				//open the file selection dialogue
                int saveFileNamed = saveFileChooser.showSaveDialog(btnOutputButton);
                String saveFilePath = "";
                if(saveFileNamed == JFileChooser.APPROVE_OPTION ) {
                	File saveFile = saveFileChooser.getSelectedFile();
                	if(!saveFile.getPath().contains(".xls") && !saveFile.getPath().contains(".xlsx")){
                		String[] inputPath = textFieldFullPathToInputFile.getText().split("\\.");
                		String extension = inputPath[inputPath.length-1];
                		
                		
                		saveFilePath = saveFileChooser.getSelectedFile().getPath() + "." + extension;
                	} else {
                		saveFilePath = saveFile.getPath();
                	}

                	if(!saveFilePath.contains(".xls") && !saveFilePath.contains(".xlsx")){
                		lblErrorMessage.setText("Please select an input file first.");
    					lblErrorMessage.setVisible(true);
    					return;
                	}
                	
                	btnOutputButton.setVisible(false);
                	textFieldFullPathToOutputFile.setText(saveFilePath);
                	textFieldFullPathToOutputFile.setVisible(true);
                }

			}
		});
	}
	
}
