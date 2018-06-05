package de.freibaer.excel;
import java.awt.BorderLayout;
import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.util.ArrayList;
import java.util.stream.Stream;

import javax.swing.DefaultListModel;
import javax.swing.JButton;
import javax.swing.JComponent;
import javax.swing.JFrame;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.ListSelectionModel;
import javax.swing.SwingUtilities;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;



public class AppFrame extends JFrame {
	private ArrayList<String> namesList = new ArrayList<>();
	
	private JTextField filterField;
	private JList<String> dataList;
	private DefaultListModel<String> listDataModel;
	
	public AppFrame() {
		//Add a stream of names as initial data
		Stream.of("Jeremy Clarkson", "James May", "Richard Hammond").forEach(namesList::add);
		constructLayout();
		setDefaultCloseOperation(EXIT_ON_CLOSE);
		setSize(500, 500);
		setVisible(true);
	}
	
	private void constructLayout() {
		add(createNorth(), BorderLayout.NORTH);
		add(createCenter(), BorderLayout.CENTER);
		add(createSouth(), BorderLayout.SOUTH);
	}
	
	private JComponent createNorth() {
		filterField = new JTextField();
		//Update the list data when the filter changes
		//We can use lambdas because of our FunctionalDocumentListener wrapper
		filterField.getDocument().addDocumentListener((FunctionalDocumentListener) (e) -> updateListData());
		return filterField;
	}
	
	private JComponent createCenter() {
		listDataModel = new DefaultListModel<>();
		namesList.stream().forEach(listDataModel::addElement); //Add initial data to the list model
		dataList = new JList<>(listDataModel);
		dataList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
		return dataList;
	}
	
	private JComponent createSouth() {
		JPanel container = new JPanel(new FlowLayout(FlowLayout.RIGHT));
		
		JButton addBtn = new JButton("Add");
		addBtn.addActionListener(this::onAddClicked); //Method reference
		container.add(addBtn);
		
		JButton removeBtn = new JButton("Remove");
		removeBtn.addActionListener(this::onRemoveClicked); //Method reference
		container.add(removeBtn);
		return container;
	}
	
	private void onAddClicked(ActionEvent e) {
		String name = JOptionPane.showInputDialog("Enter name of person");
		
		if (name != null) {
			namesList.add(name);
			updateListData();
		}
	}
	
	private void onRemoveClicked(ActionEvent e) {
		String selected = dataList.getSelectedValue();
		
		if (selected != null) {
			//Remove only if the instances are the same
			namesList.removeIf((name) -> name == selected);
			updateListData();
		}
	}
	
	private void updateListData() {
		listDataModel.clear();
		String filterText = filterField.getText().toLowerCase();
		
		//Filter away non matching names
		namesList.stream()
			.filter((name) -> filterText.isEmpty() || name.toLowerCase().contains(filterText))
			.forEach(listDataModel::addElement);;
	}
	
	/**
	 *	A wrapper around the DocumentListener interface
	 *	Allows lambdas and method reference to be used when the document changes
	 */
	@FunctionalInterface
	public interface FunctionalDocumentListener extends DocumentListener {
		@Override
		default void changedUpdate(DocumentEvent e) { onDocumentChanged(e); }
		@Override
		default void insertUpdate(DocumentEvent e) 	{ onDocumentChanged(e); }
		@Override
		default void removeUpdate(DocumentEvent e) 	{ onDocumentChanged(e); }
		
		public void onDocumentChanged(DocumentEvent e);
	}
	
	public static void main(String[] args) {
		SwingUtilities.invokeLater(AppFrame::new);
	}
}
