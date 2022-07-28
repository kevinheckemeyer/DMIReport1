import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

public class DMIReportWriter {

    public void writeToExcel(List<List<String>> records){
        String excelFilePath = "src/main/resources/DMI_Upload_template.xlsx";
        String dateTime = LocalDateTime.now().toString();
        String excelExportFilePath = "/Users/georgepeter/Downloads/DMI_Upload_template"+dateTime+".xlsx";
        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            int excelRowCount = sheet.getLastRowNum();
            int csvRow = 0;
            for(List<String> rowData : records) {
                //dont process the first 4 rows
                if(csvRow >3 && rowData.size() > 1) {
                    Row row = sheet.createRow(++excelRowCount);
                    writeRow(workbook, row, rowData);
                  }
                csvRow++;
            }
            inputStream.close();
            FileOutputStream outputStream = new FileOutputStream(excelExportFilePath);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

        } catch (IOException | EncryptedDocumentException
                 | InvalidFormatException ex) {
            ex.printStackTrace();
        }
    }
    private void writeRow(Workbook workbook,  Row row,List<String> rowData){
        CellStyle textStyle = workbook.createCellStyle();
        textStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("text"));
        Cell cell = null;
        for (int columnCount=0;columnCount<230;columnCount++) {
            cell = row.createCell(columnCount);
            if(columnCount == DMIConstant.DMIExcelCol.M03_Lender_Number.type()){
                String field1 = rowData.get(0).replace("\"", "");
                cell.setCellValue(field1);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M04_Investor_Number.type()){
                cell.setCellValue(401);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M05_Category.type()){
                cell.setCellValue("001");
                cell.setCellStyle(textStyle);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M07_Hi_Type.type()){
                cell.setCellValue(1);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M08_Loan_Type.type()){
                String field = rowData.get(2).replace("\"", "");
                cell.setCellValue(DMIConstant.getLQBLoanType(field));
            }
            if(columnCount == DMIConstant.DMIExcelCol.M10_Mortgagor_First_Name.type()){
                String field = rowData.get(4).replace("\"", "");
                field += " "+ rowData.get(5).replace("\"", "");
                cell.setCellValue(field);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M11_Mortgagor_Last_Name.type()){
                String field = rowData.get(6).replace("\"", "");
                cell.setCellValue(field);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M12_Co_Mortgagor_First_Name.type()){
                String field = rowData.get(7).replace("\"", "");
                field += " "+ rowData.get(8).replace("\"", "");
                cell.setCellValue(field);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M13_Co_Mortgagor_Last_Name.type()){
                String field = rowData.get(9).replace("\"", "");
                cell.setCellValue(field);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M14_00_Property_Address_Line_1.type()){
                String field = rowData.get(10).replace("\"", "");
                cell.setCellValue(field);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M16_Property_City.type()){
                String field = rowData.get(11).replace("\"", "");
                cell.setCellValue(field);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M17_Property_State.type()){
                String field = rowData.get(12).replace("\"", "");
                cell.setCellValue(field);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M18_Property_Zip_Code.type()){
                String field = rowData.get(13).replace("\"", "");
                cell.setCellValue(Integer.parseInt(field));
            }
            if(columnCount == DMIConstant.DMIExcelCol.M19_Mortgagor_Phone_No.type()){
                String field = rowData.get(14).replace("\"", "");
                field = field.replace("-", "");
                field = field.replace("(", "");
                field = field.replace(")", "");
                field = field.replace(" ", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.M20_Business_Number.type()){
                String field = rowData.get(15).replace("\"", "");
                field = field.replace("-", "");
                field = field.replace("(", "");
                field = field.replace(")", "");
                field = field.replace(" ", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.M23_04_Mailing_Address_Line_5.type()){
                String field = rowData.get(16).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.M24_Mailing_Address_City.type()){
                String field = rowData.get(17).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.M24_Mailing_Address_City.type()){
                String field = rowData.get(17).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.M25_Mailing_State.type()){
                String field = rowData.get(18).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.M26_00_Mailing_Zip_Code.type()){
                String field = rowData.get(19).replace("\"", "");
                cell.setCellValue(Integer.parseInt(field));
            }
            if(columnCount == DMIConstant.DMIExcelCol.M27_Mortgagor_SSN_TIN.type()){
                String field = rowData.get(20).replace("\"", "");
                cell.setCellValue(field);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M27_01_Mtg_SSN_TIN_code.type()){
                   cell.setCellValue(1);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M28_Co_Mortgagor_SSN_TIN.type()){
                String field = rowData.get(21).replace("\"", "");
                cell.setCellValue(field);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M28_01_Co_Mtg_SSN_TIN_Code.type()){
                cell.setCellValue(1);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M30_01_Original_Occupancy_Status.type()){
                String field = rowData.get(22).replace("\"", "");
                cell.setCellValue(DMIConstant.getLQBOccupancyStatusType(field.trim()));
            }
            if(columnCount == DMIConstant.DMIExcelCol.M31_Property_Construction.type()){
                cell.setCellValue("E");
            }
            if(columnCount == DMIConstant.DMIExcelCol.M32_00_Prepayment_Penalty.type()){
                String field = rowData.get(23).replace("\"", "");
                cell.setCellValue(DMIConstant.getLQBPrepaymentPenaltyType(field.trim()));
            }
            if(columnCount == DMIConstant.DMIExcelCol.M35_Original_Principal_Balance.type()){
                String field = rowData.get(24).replace("\"", "");
                cell.setCellValue(field.trim());

            }
            if(columnCount == DMIConstant.DMIExcelCol.M36_Original_Term.type()){
                String field = rowData.get(25).replace("\"", "");
                cell.setCellValue(Integer.parseInt(field));
            }
            if(columnCount == DMIConstant.DMIExcelCol.Closing_Settlement_Date.type()){
                String field = rowData.get(26).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.M40_Maturity_Date.type()){
                String field = rowData.get(26).replace("\"", "");
                if(!field.isEmpty()) {
                    String dateFormat = DMIUtils.determineDateFormat(field);
                    if (dateFormat!=null) {
                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(dateFormat);
                        //convert String to LocalDate
                        LocalDate localDate = LocalDate.parse(field, formatter);
                        String term = rowData.get(25).replace("\"", "");
                        if (!term.isEmpty()) {
                            localDate = localDate.plusMonths(Integer.parseInt(term));
                            DateTimeFormatter formatter1 = DateTimeFormatter.ofPattern("MM/dd/yyyy");
                            cell.setCellValue(localDate.format(formatter1));
                        }
                    }
                }
            }
            if(columnCount == DMIConstant.DMIExcelCol.M42_First_Payment_Date.type()){
                String field = rowData.get(27).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.M44_Payment_Frequency.type()){
                 cell.setCellValue(12);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M60_Late_Charge_Factor.type()){
                //cell.setCellValue("");
            }
            if(columnCount == DMIConstant.DMIExcelCol.M60_002_Late_Charge_Code.type()){
                cell.setCellValue("");
            }
            if(columnCount == DMIConstant.DMIExcelCol.M60_003_Late_Charge_Min_Amt.type()){
                cell.setCellValue("A");
            }
            if(columnCount == DMIConstant.DMIExcelCol.M60_004_Late_Charge_Max_Amt.type()){
                //cell.setCellValue("");
            }
            if(columnCount == DMIConstant.DMIExcelCol.M61_Grace_Days.type()){
                cell.setCellValue("15 Days");
            }
            //to-do
            if(columnCount == DMIConstant.DMIExcelCol.M62_Points_Paid_by_Borrower.type()){
                //cell.setCellValue("");
            }
            if(columnCount == DMIConstant.DMIExcelCol.M64_Investor_Loan_Number.type()){
                //cell.setCellValue("");
            }
            if(columnCount == DMIConstant.DMIExcelCol.M66_Original_Property_Value.type()){
                String field = rowData.get(29).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.M67_Purchase_Price.type()){
                String field = rowData.get(30).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.M67_Purchase_Price.type()){
                String field = rowData.get(30).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.M68_Builder_Broker_Code.type()){

            }
            if(columnCount == DMIConstant.DMIExcelCol.M69_Product_Line_Code.type()){
            }
            //need to add property type to the report instead of structure type
            if(columnCount == DMIConstant.DMIExcelCol.M72_Property_Type.type()){
                String field = rowData.get(31).replace("\"", "");
                cell.setCellValue(DMIConstant.getLQBPropertyType(field));
            }
            if(columnCount == DMIConstant.DMIExcelCol.M73_Loan_Purpose.type()){
                String field = rowData.get(32).replace("\"", "");
                cell.setCellValue(DMIConstant.getLQBLoanPurposeType(field));
            }
            if(columnCount == DMIConstant.DMIExcelCol.Owner_Type.type()){
                 cell.setCellValue(1);
            }
            if(columnCount == DMIConstant.DMIExcelCol.M75_Dist_Type.type()){
                 cell.setCellValue(1);
            }
            if(columnCount == DMIConstant.DMIExcelCol.F04_Current_Principal_Balance.type()){
                String field = rowData.get(34).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.F05_Interest_Rate.type()){
                String field = rowData.get(35).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.F06_Next_Payment_Due_Date.type()){
                String field = rowData.get(27).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.F08_P_I_Amount_Monthly.type()){
                String field = rowData.get(36).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.F08_002_County_Tax.type()){
                String field = rowData.get(68).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            //to-do need to add to report
            if(columnCount == DMIConstant.DMIExcelCol.F08_003_City_Tax.type()){

            }

            if(columnCount == DMIConstant.DMIExcelCol.F08_004_Hazard_Premium.type()){
                String field = rowData.get(69).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.F08_005_MIP.type()){
                String field = rowData.get(40).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.F08_006_Lien.type()){

                cell.setCellValue(0);
            }
            if(columnCount == DMIConstant.DMIExcelCol.F09_Escrow_Monthly_Amount.type()){

                String field = rowData.get(44).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.F13_Total_Payment.type()){
                String field1 = rowData.get(36).replace("\"", "");
                field1 = field1.replace("$", "");
                String field2 = rowData.get(68).replace("\"", "");
                field2 = field2.replace("$", "");
                String field3 = rowData.get(69).replace("\"", "");
                field3 = field3.replace("$", "");
                String field4 = rowData.get(40).replace("\"", "");
                field4 = field4.replace("$", "");
                Float totalPayment = Float.parseFloat(field1)+Float.parseFloat(field2)+Float.parseFloat(field3)+Float.parseFloat(field4);
                cell.setCellValue(totalPayment);
            }
            if(columnCount == DMIConstant.DMIExcelCol.F15_Interest_Collected_at_Closing.type()){
                String field = rowData.get(43).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            //not available in the report
            if(columnCount == DMIConstant.DMIExcelCol.F16_Escrow_Balance.type()){
                cell.setCellValue("#Manual");
            }
            if(columnCount == DMIConstant.DMIExcelCol.F20_Accrued_Late_Charges.type()){

            }
            if(columnCount == DMIConstant.DMIExcelCol.F21_Late_Charges_Paid_YTD.type()){

            }
            if(columnCount == DMIConstant.DMIExcelCol.F22_Interest_Paid_YTD.type()){
                String field = rowData.get(43).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.F23_Principal_Paid_YTD.type()){
                cell.setCellValue("#Manual");
            }
            if(columnCount == DMIConstant.DMIExcelCol.F24_Taxes_Paid_YTD.type()){
                cell.setCellValue("#Manual");
            }
            if(columnCount == DMIConstant.DMIExcelCol.F29_IOE_Paid_YTD.type()){

            }
            if(columnCount == DMIConstant.DMIExcelCol.F31_01_Effective_Yield.type()){

            }
            if(columnCount == DMIConstant.DMIExcelCol.F32_Origination_Fees_Costs_1.type()){

            }
            if(columnCount == DMIConstant.DMIExcelCol.F33_Origination_Fees_Costs_2.type()){

            }
            if(columnCount == DMIConstant.DMIExcelCol.F34_Origination_Fees_Costs_3.type()){

            }
            if(columnCount == DMIConstant.DMIExcelCol.Z03_Sale_Id.type()){
                cell.setCellValue("T74FLOW");
            }
            if(columnCount == DMIConstant.DMIExcelCol.Z04_Man_Code.type()){
                cell.setCellValue("U");
            }
            if(columnCount == DMIConstant.DMIExcelCol.Z05_Bill_Mode.type()){
                cell.setCellValue(9);
            }
            if(columnCount == DMIConstant.DMIExcelCol.Z14_3_Pos_Fld_3A.type()){
                cell.setCellValue("T74");
            }
            if(columnCount == DMIConstant.DMIExcelCol.Z15_Zone_Branch_Code.type()){
                cell.setCellValue("EL");
            }
            if(columnCount == DMIConstant.DMIExcelCol.Z19_Cash_Batch_Code.type()){
                cell.setCellValue("RE");
            }
            if(columnCount == DMIConstant.DMIExcelCol.Next_tax_due_date.type()){
                String field = rowData.get(45).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            //need to add to the report
            if(columnCount == DMIConstant.DMIExcelCol.Z26_Branch_Code.type()){
                 cell.setCellValue("#Manual");
            }
            if(columnCount == DMIConstant.DMIExcelCol.X03_Mers_Min.type()){
                String field = rowData.get(48).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.X04_Mers_Min_Reg_Date.type()){
                String field = rowData.get(49).replace("\"", "");
                cell.setCellValue(field.trim());
            }
            if(columnCount == DMIConstant.DMIExcelCol.X05_Mers_Min_Reg_Flag.type()){
                 cell.setCellValue("Y");
            }
            if(columnCount == DMIConstant.DMIExcelCol.X06_Mers_Mom_Flag.type()){
                cell.setCellValue("I");
            }





        }
    }

    }

