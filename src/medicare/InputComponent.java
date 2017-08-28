package medicare;

import java.awt.BorderLayout;
import java.awt.CardLayout;
import java.awt.event.ActionListener;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JPanel;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

public class InputComponent {
	public String file = "";// ����excel·��
//	public static String beginTime = "";
//	public static String endTime = "";
	public int reportNo = 0;// ����� Ĭ��0�ű���

	JFrame f = new JFrame("ҽ������----ɽ��ʡ����ҽԺ");
	// �˵���
	JMenuBar mb = new JMenuBar();
	JMenu menu = new JMenu("ҽ������");
	private CardLayout card = new CardLayout(5, 5);
	JPanel cardPanel = new JPanel(card);

	JFileChooser fc1 = new JFileChooser();

	public void init() {
		// �½�����panel,Ȼ����ص�cardPanel
		ReportPanel[] p = new ReportPanel[ReportNameConstants.report.length];
		for (int i = 0; i < ReportNameConstants.report.length; i++) {
			p[i] = new ReportPanel(ReportNameConstants.report[i]);
			cardPanel.add(p[i], String.valueOf(i));// ����card��
		}

		// �˵�������
		ActionListener al = e -> {
			for (int i = 0; i < ReportNameConstants.report.length; i++) {
				if (e.getActionCommand() == ReportNameConstants.report[i]) {
					card.show(cardPanel, String.valueOf(i));
					reportNo = i;
					break;
				}
			}
		};
		// �ļ�ѡ����������
		ActionListener fileChooser = e -> {
			fc1.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);// ֻ��ѡ·��,����ѡ�ļ�
			fc1.showOpenDialog(f);// ��ʾJFileChooser
			fc1.setFileSelectionMode(1);// �趨ֻ��ѡ���ļ���
			file = fc1.getSelectedFile().getAbsolutePath();
		};
		// ��ʼ����button������
		ActionListener beginOut = e -> {
			// ��ȡ��ʼʱ�䡢����ʱ��
			Object x = e.getSource();// ȡ���˵��ü�������button
			for (int i = 0; i < ReportNameConstants.report.length; i++) {
				if (x.equals(p[i].getComponent(6))) {
					ReportTimeConstants.beginTime = ((DateChooserJButton) p[i].getComponent(2)).getText();
					ReportTimeConstants.endTime = ((DateChooserJButton) p[i].getComponent(4)).getText();
				}
			}
			new OutputExcel(ReportTimeConstants.beginTime, ReportTimeConstants.endTime, file, reportNo);// ���õ�����
		};

		// f.setBounds(200, 200,500,500);
		// ���ز˵�
		JMenuItem[] mi = new JMenuItem[ReportNameConstants.report.length];
		for (int i = 0; i < ReportNameConstants.report.length; i++) {
			mi[i] = new JMenuItem(ReportNameConstants.report[i]);
			mi[i].addActionListener(al);
			menu.add(mi[i]);
		}
		mb.add(menu);
		f.add(mb, BorderLayout.NORTH);
		// ����cardPanel
		f.add(cardPanel);
		for (int i = 0; i < ReportNameConstants.report.length; i++) {
			p[i].out.addActionListener(fileChooser);// �����ļ�ѡ����������
			p[i].bn.addActionListener(beginOut);// ���ؿ�ʼ����button������
		}
		// ������۷��
		try {
			// javax.swing.plaf.nimbus.NimbusLookAndFeel
			// com.sun.java.swing.plaf.windows.WindowsLookAndFeel
			UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
			SwingUtilities.updateComponentTreeUI(f);
			SwingUtilities.updateComponentTreeUI(fc1);
		} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
				| UnsupportedLookAndFeelException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		f.pack();
		// ��ʾ���崰��
		f.setVisible(true);

	}

	public static void main(String[] args) {
		new InputComponent().init();
	}
}
