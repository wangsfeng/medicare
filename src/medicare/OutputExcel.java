package medicare;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import javax.swing.JOptionPane;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OutputExcel {

	public OutputExcel(String beginTime, String endTime, String file, int reportNo) {
		final String connectionUrl = "jdbc:sqlserver://111.111.116.12:1433;databaseName=zlyyhis2000;integratedSecurity=false;";
		final String user = "sa";
		final String psd = "M83T75Z67C57";
		String[] sql = new String[ReportNameConstants.report.length];		
		//0���໼���ܷ����˴�
		sql[0]= "select PatientProperty as ����������,PatientTypeName as ��������,sum(TotalTcMoney) as ͳ����� "
				+ ",sum(TotalCosts) as �ܷ���,count( DISTINCT A.No_TreatList) as �˴� " + "From AdmBalance a "
				+ "join DictPatientType b on a.PatientProperty=b.No_PatientType " + "where isnull(a.BalanceStatus,0) = 1 and BalanceDateTime>='"
				+ beginTime + "'  and BalanceDateTime<'" + endTime + "' group by PatientProperty,PatientTypeName ";
		//1���໼���˴���ͷ��
		sql[1]="SELECT count(a.No_TreatList) as �˴�,count(distinct InPCode) as ��ͷ,PatientProperty as ���������� "
				+ ",PatientTypeName as �������� FROM admpatsvisit A "
				+ "join DictPatientType b on a.PatientProperty=b.No_PatientType "
				+ "where a.DischargeDate>='"+beginTime
				+ "' AND A.DischargeDate<'"+endTime
				+ "'group by PatientProperty,PatientTypeName ";
		//2ʡҽ�������˴���ͷ��
		sql[2]="SELECT count(a.No_TreatList) as �˴�,count(distinct InPCode) as ��ͷ,PatientProperty as ҽ������ "
				+ "FROM admpatsvisit A 	join pInsInHosRecord p on p.no_TreatList=a.No_TreatList "
				+ "where PatientProperty =800 "
				+ "	AND a.DischargeDate>='"+beginTime
				+ "' AND a.DischargeDate<'"+endTime
				+ "'AND p.cInsCode <> p.cIcCard group by PatientProperty";
		//3ʡҽ�������˴���ͷ����ϸ
		sql[3]="select distinct InPCode  as סԺ��,d1.DeptName ��Ժ����,d2.DeptName ��Ժ����,count(a.No_TreatList)  as �˴���ͷ��, "
				+ "PatientProperty as ҽ������ from admpatsvisit A "
				+ "left join dictdept d1  on a.InDept=d1.No_Dept "
				+ "left join dictdept d2  on a.OutDept=d2.No_Dept "
				+ "join pInsInHosRecord p on p.no_TreatList=a.No_TreatList "
				+ "where PatientProperty =800 AND p.cInsCode <> p.cIcCard AND a.OutDept IS NOT NULL "
				+ "AND a.DischargeDate>='"+beginTime
				+ "' AND a.DischargeDate<'"+endTime
				+ "' group by PatientProperty,InPCode,d1.DeptName,d2.DeptName order by �˴���ͷ�� desc";
		//4ʡҽ�������˴���ͷ�ȷֿ���ͳ��
		sql[4]="SELECT count(a.No_TreatList) as �˴�,count(distinct InPCode) as ��ͷ,d2.DeptName ��Ժ����,PatientProperty as ҽ������ "
				+ "FROM admpatsvisit A left join dictdept d2 on a.OutDept=d2.No_Dept "
				+ "join pInsInHosRecord p on p.no_TreatList=a.No_TreatList "
				+ "where PatientProperty =800 AND a.DischargeDate>='"+beginTime
				+ "' AND a.DischargeDate<'"+endTime
				+ "' AND a.OutDept IS NOT NULL  AND p.cInsCode <> p.cIcCard group by PatientProperty,d2.DeptName";
		//ʡҽ����ط���ͳ�﹫���˴�
		sql[5]="select sum(nsumrate) as �����ܶ�,SUM( nworldrate) ͳ�� ,SUM(nofficialrate) ����,SUM( nworldrate) + SUM(nofficialrate) ����ܶ�,count(p.no_TreatList) as �˴� "
				+ "from  pinssquareinfor p join    AdmBalance a on p.no_TreatList = a.No_TreatList "
				+ "join pInsInHosRecord p2  on p.no_TreatList=p2.No_TreatList where a.BalanceDateTime >= '"+beginTime
				+ "' and a.BalanceDateTime < '"+endTime
				+ "' and isnull(p.bValid,0) = 1  and isnull(a.BalanceStatus,0) = 1 	AND p2.cInsCode = p2.cIcCard";
		sql[6]="select sum(nsumrate) as �����ܶ�,SUM( nworldrate) ͳ�� ,SUM(nofficialrate) ����,SUM( nworldrate) + SUM(nofficialrate) ����ܶ�,count(p.no_TreatList) as �˴� "
				+ "from  pinssquareinfor p join    AdmBalance a on p.no_TreatList = a.No_TreatList "
				+ "join pInsInHosRecord p2  on p.no_TreatList=p2.No_TreatList where a.BalanceDateTime >= '"+beginTime
				+ "' and a.BalanceDateTime < '"+endTime
				+ "' and isnull(p.bValid,0) = 1 and isnull(a.BalanceStatus,0) = 1 AND p2.cInsCode <> p2.cIcCard";

		sql[7]="select year(BalanceDateTime) as '��',MONTH(BalanceDateTime) as '��',(CASE when p2.cInsCode = p2.cIcCard  THEN '���' else '����' end) as '�������', "
				+ "sum(nsumrate) as �����ܶ�,SUM( nworldrate) ͳ�� ,SUM(nofficialrate) ����,SUM( nworldrate) + SUM(nofficialrate) ����ܶ�,count(p.no_TreatList) as �˴� "
				+ "from  pinssquareinfor p join    AdmBalance a  on p.no_TreatList = a.No_TreatList "
				+ "join pInsInHosRecord p2 on p.no_TreatList=p2.No_TreatList where BalanceDateTime>='"+beginTime
				+ "' and BalanceDateTime<'"+endTime
				+ "'AND isnull(p.bValid,0) = 1 and isnull(a.BalanceStatus,0) = 1 "
				+ "GROUP BY  year(BalanceDateTime),MONTH(BalanceDateTime),(CASE when p2.cInsCode = p2.cIcCard  THEN '���' else '����' end) "
				+ "order by  year(BalanceDateTime),MONTH(BalanceDateTime),(CASE when p2.cInsCode = p2.cIcCard  THEN '���' else '����' end) ";
		//����ʡҽ����ϸ
		sql[8]="select nsumrate as �����ܶ�, nworldrate ͳ�� ,nofficialrate ����, nworldrate + nofficialrate ����ܶ�,cPersonName ����,p.no_TreatList �����,BalanceDateTime �������� "
				+"from  pinssquareinfor p join    AdmBalance a  on p.no_TreatList = a.No_TreatList "
				+"join pInsInHosRecord p2  on p.no_TreatList=p2.No_TreatList where a.BalanceDateTime >= '"+beginTime
				+ "' and a.BalanceDateTime < '"+endTime
				+ "' and isnull(p.bValid,0) = 1  and isnull(a.BalanceStatus,0) = 1 AND p2.cInsCode <> p2.cIcCard ";
		sql[9]="select d.DeptName ��Ժ���� ,(CASE when p2.cInsCode = p2.cIcCard  THEN '���' else '����' end) as '�������', "
				+"sum(nsumrate) as �����ܶ�,SUM( nworldrate) ͳ�� ,SUM(nofficialrate) ����,SUM( nworldrate) + SUM(nofficialrate) ����ܶ�,count(p.no_TreatList) as �˴� "
				+ "from  pinssquareinfor p join    AdmBalance a  on p.no_TreatList = a.No_TreatList "
				+ "join pInsInHosRecord p2 on p.no_TreatList=p2.No_TreatList "
				+ "join dictdept d on d.No_Dept=a.BalanceDept "
				+ "where a.BalanceDateTime >= '"+beginTime
				+ "' and a.BalanceDateTime < '"+endTime
				+ "' and isnull(p.bValid,0) = 1 and isnull(a.BalanceStatus,0) = 1 "
				+ "group by (CASE when p2.cInsCode = p2.cIcCard  THEN '���' else '����' end),d.DeptName "
				+ "order by d.DeptName,(CASE when p2.cInsCode = p2.cIcCard  THEN '���' else '����' end) ";
		sql[10]="select sum(nMoney) �ܷ���,sum(nselfmoney) ͳ��,count(iDiagnoseCode) �˴�,isnull(imedicalflag,0) ҽ��������,  "
				+ "(case when imedicalflag=1 then '��ҽ��' when imedicalflag=2 then 'ʡҽ��' when imedicalflag=12 then '��·ҽ��' "
				+ "when imedicalflag=-4 then '�����ɲ�'  else '�Է�' end) as ҽ������  from opdratemain "
				+ "where dgetratedate>='"+beginTime
				+ "' and dgetratedate<'"+endTime
				+ "' group by imedicalflag ";
		sql[11]="select isnull(a.iWholePay,0) as ͳ���˻�,a.No_TreatList as �����,b.InPCode as סԺ��, PatName as ����,d.DeptName as  ���� "
				+ ", b.PatientType ����������, c.PatientTypeName �������� ,  "
				+ "case when p.cInsCode <> p.cIcCard then '����' when p.cInsCode=p.cIcCard then '���' else '' end   ʡҽ��������� "
				+ "   from AdmPatsInHospital a join AdmPatsVisit b on a.No_TreatList=b.No_TreatList "
				+ "  left join pInsInHosRecord p on a.No_TreatList=p.No_TreatList "
				+ "   left join dictdept d on d.no_dept =a.No_ClinicManDept "
				+ "   left join  DictPatientType c on b.PatientType=c.No_PatientType";
		sql[12]="select count(*) as �˴�,sum(nMoney) ���� from  opdratemain where dgetratedate>='"+beginTime
				+ "' and dgetratedate<'"+endTime
				+ "' and imedicalflag=1 and csickcode is not  null ";
		sql[13]="USE OLDMR "
				+ "SELECT d.DeptName as ��������,count(a.SickId) as �����˴�,sum(a.TotalCosts) as סԺ����,p.PatientTypeName as ����ҽ����� "
				+ "FROM PATVISIT A join zlyyhis2000.dbo.admpatsvisit c on (a.VisitId=c.No_Visit and a.CaseCode=c.CaseCode) "
				+ "join zlyyhis2000.dbo.dictdept d on d.No_Dept=c.OutDept "
				+ "join zlyyhis2000.dbo.[DictPatientType] p on  p.No_PatientType=c.PatientType "
				+ "where exists (select * from  Operation o  where A.SickId=o.SickId and A.VisitId=o.VisitId) "
				+ "and  a.OutDateTime>='"+beginTime
				+ "' AND a.OutDateTime<'"+endTime
				+ "' and c.OutDept is not null  group by d.DeptName,p.PatientTypeName order by d.DeptName";
		sql[14]="USE OLDMR "
				+ "SELECT d.DeptName as ��������,count(a.SickId) as ���˴�,sum(a.TotalCosts) as סԺ����,p.PatientTypeName as ����ҽ����� "
				+ "FROM PATVISIT A join zlyyhis2000.dbo.admpatsvisit c  on (a.VisitId=c.No_Visit and a.CaseCode=c.CaseCode) "
				+ "join zlyyhis2000.dbo.dictdept d on d.No_Dept=c.OutDept "
				+ "join zlyyhis2000.dbo.[DictPatientType] p on  p.No_PatientType=c.PatientType "
				+ "where  a.OutDateTime>='"+beginTime
				+ "' AND a.OutDateTime<'"+endTime
				+ "' and c.OutDept is not null group by d.DeptName,p.PatientTypeName order by d.DeptName ";

				
				
		
		
		
		try {
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
			Connection con = DriverManager.getConnection(connectionUrl, user, psd);
			Statement stmt = con.createStatement();
			
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet();
			// Ϊ�����ڵ�Ԫ����д�����ģ������ַ�����ΪUTF_16��
			// workbook.setSheetName(0,sheetName,XSSFWorkbook.ENCODING_UTF_16);
			XSSFRow row = sheet.createRow((short) 0);
			;
			XSSFCell cell;
			ResultSet rs = stmt.executeQuery(sql[reportNo]);
			ResultSetMetaData md = rs.getMetaData();
			int nColumn = md.getColumnCount();
			// д������ֶε�����
			for (int i = 1; i <= nColumn; i++) {
				cell = row.createCell((short) (i - 1));
				cell.setCellType(XSSFCell.CELL_TYPE_STRING);
				// ��������
				// cell.setEncoding(XSSFCell.ENCODING_UTF_16);
				cell.setCellValue(md.getColumnLabel(i));
			}
			int iRow = 1;
			// д�������¼��ÿ����¼��ӦExcel�е�һ��
			while (rs.next()) {
				row = sheet.createRow((short) iRow);
				for (int j = 1; j <= nColumn; j++) {
					cell = row.createCell((short) (j - 1));
					cell.setCellType(XSSFCell.CELL_TYPE_STRING);
					// cell.setEncoding(XSSFCell.ENCODING_UTF_16);
					cell.setCellValue(rs.getObject(j).toString());
				}
				iRow++;
			}
			FileOutputStream fOut = new FileOutputStream(file + "\\����" + reportNo + ".xlsx");
			workbook.write(fOut);
			fOut.flush();
			fOut.close();
//			System.out.println("�������ݳɹ���");
			JOptionPane.showMessageDialog(null,"�������ݳɹ���");

		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
