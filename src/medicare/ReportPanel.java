package medicare;

import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.JPanel;

public class ReportPanel extends JPanel {

	// ���ݱ����������
	public JLabel begin = new JLabel("��ʼʱ��");
	public DateChooserJButton beginTime = new DateChooserJButton(ReportTimeConstants.beginTime);
	public JLabel end = new JLabel("����ʱ��");
	public DateChooserJButton endTime = new DateChooserJButton(ReportTimeConstants.endTime);
	public JButton out = new JButton("ѡ�񵼳�·��");
	public JButton bn = new JButton("��ʼ����");

	// ���캯��
	public ReportPanel(String arg) {
		super();
		JLabel name = new JLabel(arg);
		// �������ݱ����������
		this.add(name);
		this.add(begin);
		this.add(beginTime);
		this.add(end);
		this.add(endTime);
		this.add(out);
		this.add(bn);
	}
}
