import java.awt.BorderLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.SwingUtilities;


@SuppressWarnings("serial")
public class GTAGUI extends JPanel implements ActionListener{
	//fields
		static private final String newline = "\n";
		static File inputPath;
	    JButton chooseFolderButton, runButton;
	    static JTextArea log;
	    JFileChooser fc;

	//constructor
		public GTAGUI () {
			super (new BorderLayout());
			
			//Create the log first, because the action listeners
	        log = new JTextArea(25,60);
	        log.setMargin(new Insets(5,5,5,5));
	        log.setEditable(false);
	        JScrollPane logScrollPane = new JScrollPane(log);
	        
	        //Create a file chooser
	        fc = new JFileChooser();
	        
	        //Create the choose folder button
	        chooseFolderButton = new JButton ("Select Source Folder...");
	        chooseFolderButton.addActionListener(this);
	        
	        //Create the choose folder button
	        runButton = new JButton ("Run");
	        runButton.addActionListener(this);
	        
	        //For layout purposes, put the buttons in a separate panel
	        JPanel buttonPanel = new JPanel(); //use FlowLayout
	        buttonPanel.add(chooseFolderButton);
	        buttonPanel.add(runButton);

	        //Add the buttons and the log to this panel.
	        add(buttonPanel, BorderLayout.PAGE_START);
	        add(logScrollPane, BorderLayout.CENTER);
		}

	@Override
	public void actionPerformed(ActionEvent ev) {
		//Handle choose folder button action.
        if (ev.getSource() == chooseFolderButton) {
        	fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            int returnVal = fc.showOpenDialog(GTAGUI.this);

            if (returnVal == JFileChooser.APPROVE_OPTION) {
                inputPath = fc.getSelectedFile();
                if (!inputPath.getName().equalsIgnoreCase("grid")) {
                	log.append("Please select folder </Users/username/Desktop/grid/>" + newline);
                	inputPath = null;
                } else { 
                	log.append("Selecting Source Folder: " + inputPath.getName() +"/" + newline);
                }
            } else {
                log.append("Select Source Foulder cancelled by user." + newline);
            }
            log.setCaretPosition(log.getDocument().getLength());
            
        //Handle run button action.   
        } else if (ev.getSource() == runButton) {
        	if (inputPath != null) {
        		if (inputPath.list().length < 2) {
        			log.append("Source Folder is empty!" + newline);	
        		} else {
        			try {
        				Thread t = new Thread( new Runnable () {
        						public void run() {
        							GTAFunctions.countAndCheck(inputPath);
        						}
        				});
        				t.start();
        			} catch (Exception e) {
        				log.append("Exception! Unable to complete." + newline);
        			} finally {
        				//do nothing
        			}
        		}

        	} else {
            	log.append("Select Source Folder first." + newline);
            }
        	log.setCaretPosition(log.getDocument().getLength());
        }
	}
	
	private static void createAndShowGUI () {
		//create the window
		JFrame frame = new JFrame ("GRID To Arrow - Forecast Template Transfer");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		//add content
		frame.add(new GTAGUI());
		
		//display the window
		frame.pack();
		frame.setVisible(true);
		
	}

	public static void generalMessage (String msg) {
		log.append(msg + newline);
		log.setCaretPosition(log.getDocument().getLength());
	}
	
	public static void main(String[] args) {
		SwingUtilities.invokeLater(new Runnable() {
			public void run () {
				createAndShowGUI();
			}
		});

	}
}
