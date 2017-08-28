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
		//0各类患者总费用人次
		sql[0]= "select PatientProperty as 患者类型码,PatientTypeName as 患者类型,sum(TotalTcMoney) as 统筹费用 "
				+ ",sum(TotalCosts) as 总费用,count( DISTINCT A.No_TreatList) as 人次 " + "From AdmBalance a "
				+ "join DictPatientType b on a.PatientProperty=b.No_PatientType " + "where isnull(a.BalanceStatus,0) = 1 and BalanceDateTime>='"
				+ beginTime + "'  and BalanceDateTime<'" + endTime + "' group by PatientProperty,PatientTypeName ";
		//1各类患者人次人头比
		sql[1]="SELECT count(a.No_TreatList) as 人次,count(distinct InPCode) as 人头,PatientProperty as 患者类型码 "
				+ ",PatientTypeName as 患者类型 FROM admpatsvisit A "
				+ "join DictPatientType b on a.PatientProperty=b.No_PatientType "
				+ "where a.DischargeDate>='"+beginTime
				+ "' AND A.DischargeDate<'"+endTime
				+ "'group by PatientProperty,PatientTypeName ";
		//2省医保本地人次人头比
		sql[2]="SELECT count(a.No_TreatList) as 人次,count(distinct InPCode) as 人头,PatientProperty as 医保类型 "
				+ "FROM admpatsvisit A 	join pInsInHosRecord p on p.no_TreatList=a.No_TreatList "
				+ "where PatientProperty =800 "
				+ "	AND a.DischargeDate>='"+beginTime
				+ "' AND a.DischargeDate<'"+endTime
				+ "'AND p.cInsCode <> p.cIcCard group by PatientProperty";
		//3省医保本地人次人头比明细
		sql[3]="select distinct InPCode  as 住院号,d1.DeptName 入院科室,d2.DeptName 出院科室,count(a.No_TreatList)  as 人次人头比, "
				+ "PatientProperty as 医保类型 from admpatsvisit A "
				+ "left join dictdept d1  on a.InDept=d1.No_Dept "
				+ "left join dictdept d2  on a.OutDept=d2.No_Dept "
				+ "join pInsInHosRecord p on p.no_TreatList=a.No_TreatList "
				+ "where PatientProperty =800 AND p.cInsCode <> p.cIcCard AND a.OutDept IS NOT NULL "
				+ "AND a.DischargeDate>='"+beginTime
				+ "' AND a.DischargeDate<'"+endTime
				+ "' group by PatientProperty,InPCode,d1.DeptName,d2.DeptName order by 人次人头比 desc";
		//4省医保本地人次人头比分科室统计
		sql[4]="SELECT count(a.No_TreatList) as 人次,count(distinct InPCode) as 人头,d2.DeptName 出院科室,PatientProperty as 医保类型 "
				+ "FROM admpatsvisit A left join dictdept d2 on a.OutDept=d2.No_Dept "
				+ "join pInsInHosRecord p on p.no_TreatList=a.No_TreatList "
				+ "where PatientProperty =800 AND a.DischargeDate>='"+beginTime
				+ "' AND a.DischargeDate<'"+endTime
				+ "' AND a.OutDept IS NOT NULL  AND p.cInsCode <> p.cIcCard group by PatientProperty,d2.DeptName";
		//省医保异地费用统筹公补人次
		sql[5]="select sum(nsumrate) as 费用总额,SUM( nworldrate) 统筹 ,SUM(nofficialrate) 公补,SUM( nworldrate) + SUM(nofficialrate) 相加总额,count(p.no_TreatList) as 人次 "
				+ "from  pinssquareinfor p join    AdmBalance a on p.no_TreatList = a.No_TreatList "
				+ "join pInsInHosRecord p2  on p.no_TreatList=p2.No_TreatList where a.BalanceDateTime >= '"+beginTime
				+ "' and a.BalanceDateTime < '"+endTime
				+ "' and isnull(p.bValid,0) = 1  and isnull(a.BalanceStatus,0) = 1 	AND p2.cInsCode = p2.cIcCard";
		sql[6]="select sum(nsumrate) as 费用总额,SUM( nworldrate) 统筹 ,SUM(nofficialrate) 公补,SUM( nworldrate) + SUM(nofficialrate) 相加总额,count(p.no_TreatList) as 人次 "
				+ "from  pinssquareinfor p join    AdmBalance a on p.no_TreatList = a.No_TreatList "
				+ "join pInsInHosRecord p2  on p.no_TreatList=p2.No_TreatList where a.BalanceDateTime >= '"+beginTime
				+ "' and a.BalanceDateTime < '"+endTime
				+ "' and isnull(p.bValid,0) = 1 and isnull(a.BalanceStatus,0) = 1 AND p2.cInsCode <> p2.cIcCard";

		sql[7]="select year(BalanceDateTime) as '年',MONTH(BalanceDateTime) as '月',(CASE when p2.cInsCode = p2.cIcCard  THEN '异地' else '本地' end) as '本地异地', "
				+ "sum(nsumrate) as 费用总额,SUM( nworldrate) 统筹 ,SUM(nofficialrate) 公补,SUM( nworldrate) + SUM(nofficialrate) 相加总额,count(p.no_TreatList) as 人次 "
				+ "from  pinssquareinfor p join    AdmBalance a  on p.no_TreatList = a.No_TreatList "
				+ "join pInsInHosRecord p2 on p.no_TreatList=p2.No_TreatList where BalanceDateTime>='"+beginTime
				+ "' and BalanceDateTime<'"+endTime
				+ "'AND isnull(p.bValid,0) = 1 and isnull(a.BalanceStatus,0) = 1 "
				+ "GROUP BY  year(BalanceDateTime),MONTH(BalanceDateTime),(CASE when p2.cInsCode = p2.cIcCard  THEN '异地' else '本地' end) "
				+ "order by  year(BalanceDateTime),MONTH(BalanceDateTime),(CASE when p2.cInsCode = p2.cIcCard  THEN '异地' else '本地' end) ";
		//本地省医保明细
		sql[8]="select nsumrate as 费用总额, nworldrate 统筹 ,nofficialrate 公补, nworldrate + nofficialrate 相加总额,cPersonName 姓名,p.no_TreatList 就诊号,BalanceDateTime 结算日期 "
				+"from  pinssquareinfor p join    AdmBalance a  on p.no_TreatList = a.No_TreatList "
				+"join pInsInHosRecord p2  on p.no_TreatList=p2.No_TreatList where a.BalanceDateTime >= '"+beginTime
				+ "' and a.BalanceDateTime < '"+endTime
				+ "' and isnull(p.bValid,0) = 1  and isnull(a.BalanceStatus,0) = 1 AND p2.cInsCode <> p2.cIcCard ";
		sql[9]="select d.DeptName 出院科室 ,(CASE when p2.cInsCode = p2.cIcCard  THEN '异地' else '本地' end) as '本地异地', "
				+"sum(nsumrate) as 费用总额,SUM( nworldrate) 统筹 ,SUM(nofficialrate) 公补,SUM( nworldrate) + SUM(nofficialrate) 相加总额,count(p.no_TreatList) as 人次 "
				+ "from  pinssquareinfor p join    AdmBalance a  on p.no_TreatList = a.No_TreatList "
				+ "join pInsInHosRecord p2 on p.no_TreatList=p2.No_TreatList "
				+ "join dictdept d on d.No_Dept=a.BalanceDept "
				+ "where a.BalanceDateTime >= '"+beginTime
				+ "' and a.BalanceDateTime < '"+endTime
				+ "' and isnull(p.bValid,0) = 1 and isnull(a.BalanceStatus,0) = 1 "
				+ "group by (CASE when p2.cInsCode = p2.cIcCard  THEN '异地' else '本地' end),d.DeptName "
				+ "order by d.DeptName,(CASE when p2.cInsCode = p2.cIcCard  THEN '异地' else '本地' end) ";
		sql[10]="select sum(nMoney) 总费用,sum(nselfmoney) 统筹,count(iDiagnoseCode) 人次,isnull(imedicalflag,0) 医保类型码,  "
				+ "(case when imedicalflag=1 then '市医保' when imedicalflag=2 then '省医保' when imedicalflag=12 then '铁路医保' "
				+ "when imedicalflag=-4 then '保健干部'  else '自费' end) as 医保类型  from opdratemain "
				+ "where dgetratedate>='"+beginTime
				+ "' and dgetratedate<'"+endTime
				+ "' group by imedicalflag ";
		sql[11]="select isnull(a.iWholePay,0) as 统筹账户,a.No_TreatList as 就诊号,b.InPCode as 住院号, PatName as 姓名,d.DeptName as  科室 "
				+ ", b.PatientType 患者类型码, c.PatientTypeName 患者类型 ,  "
				+ "case when p.cInsCode <> p.cIcCard then '本地' when p.cInsCode=p.cIcCard then '异地' else '' end   省医保本地异地 "
				+ "   from AdmPatsInHospital a join AdmPatsVisit b on a.No_TreatList=b.No_TreatList "
				+ "  left join pInsInHosRecord p on a.No_TreatList=p.No_TreatList "
				+ "   left join dictdept d on d.no_dept =a.No_ClinicManDept "
				+ "   left join  DictPatientType c on b.PatientType=c.No_PatientType";
		sql[12]="select count(*) as 人次,sum(nMoney) 费用 from  opdratemain where dgetratedate>='"+beginTime
				+ "' and dgetratedate<'"+endTime
				+ "' and imedicalflag=1 and csickcode is not  null ";
		sql[13]="USE OLDMR "
				+ "SELECT d.DeptName as 科室名称,count(a.SickId) as 手术人次,sum(a.TotalCosts) as 住院费用,p.PatientTypeName as 患者医保类别 "
				+ "FROM PATVISIT A join zlyyhis2000.dbo.admpatsvisit c on (a.VisitId=c.No_Visit and a.CaseCode=c.CaseCode) "
				+ "join zlyyhis2000.dbo.dictdept d on d.No_Dept=c.OutDept "
				+ "join zlyyhis2000.dbo.[DictPatientType] p on  p.No_PatientType=c.PatientType "
				+ "where exists (select * from  Operation o  where A.SickId=o.SickId and A.VisitId=o.VisitId) "
				+ "and  a.OutDateTime>='"+beginTime
				+ "' AND a.OutDateTime<'"+endTime
				+ "' and c.OutDept is not null  group by d.DeptName,p.PatientTypeName order by d.DeptName";
		sql[14]="USE OLDMR "
				+ "SELECT d.DeptName as 科室名称,count(a.SickId) as 总人次,sum(a.TotalCosts) as 住院费用,p.PatientTypeName as 患者医保类别 "
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
			// 为了能在单元格中写入中文，设置字符编码为UTF_16。
			// workbook.setSheetName(0,sheetName,XSSFWorkbook.ENCODING_UTF_16);
			XSSFRow row = sheet.createRow((short) 0);
			;
			XSSFCell cell;
			ResultSet rs = stmt.executeQuery(sql[reportNo]);
			ResultSetMetaData md = rs.getMetaData();
			int nColumn = md.getColumnCount();
			// 写入各个字段的名称
			for (int i = 1; i <= nColumn; i++) {
				cell = row.createCell((short) (i - 1));
				cell.setCellType(XSSFCell.CELL_TYPE_STRING);
				// 设置中文
				// cell.setEncoding(XSSFCell.ENCODING_UTF_16);
				cell.setCellValue(md.getColumnLabel(i));
			}
			int iRow = 1;
			// 写入各条记录，每条记录对应Excel中的一行
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
			FileOutputStream fOut = new FileOutputStream(file + "\\数据" + reportNo + ".xlsx");
			workbook.write(fOut);
			fOut.flush();
			fOut.close();
//			System.out.println("导出数据成功！");
			JOptionPane.showMessageDialog(null,"导出数据成功！");

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
